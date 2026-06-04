using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelOps;

/// <summary>
/// Utilities for parsing A1-style cell references and safely inserting
/// cells into a worksheet in sorted order (required by OOXML spec).
/// </summary>
public static class A1
{
    // ---------------------------------------------------------------------------
    // Parse / format
    // ---------------------------------------------------------------------------

    /// <summary>
    /// Parse an A1 reference (e.g. "B3") into 1-based (col, row) ints.
    /// </summary>
    public static (int Col, int Row) Parse(string cellRef)
    {
        if (string.IsNullOrWhiteSpace(cellRef))
            throw new OpException(ErrorCategory.RangeParseError, $"Empty cell reference");

        int i = 0;
        while (i < cellRef.Length && char.IsLetter(cellRef[i])) i++;
        if (i == 0 || i == cellRef.Length)
            throw new OpException(ErrorCategory.RangeParseError, $"Invalid A1 reference: {cellRef}");

        string colStr = cellRef[..i].ToUpperInvariant();
        string rowStr = cellRef[i..];

        if (!int.TryParse(rowStr, out int row) || row < 1)
            throw new OpException(ErrorCategory.RangeParseError, $"Invalid row number in: {cellRef}");

        int col = 0;
        foreach (char c in colStr)
        {
            if (c < 'A' || c > 'Z')
                throw new OpException(ErrorCategory.RangeParseError, $"Invalid column letter in: {cellRef}");
            col = col * 26 + (c - 'A' + 1);
        }

        return (col, row);
    }

    /// <summary>
    /// Convert 1-based column index to letter(s), e.g. 1→"A", 27→"AA".
    /// </summary>
    public static string ColToLetters(int col)
    {
        string result = "";
        while (col > 0)
        {
            col--;
            result = (char)('A' + col % 26) + result;
            col /= 26;
        }
        return result;
    }

    /// <summary>
    /// Convert (col, row) back to an A1 string.
    /// </summary>
    public static string ToA1(int col, int row) => $"{ColToLetters(col)}{row}";

    // ---------------------------------------------------------------------------
    // EnsureCell — insert-in-order helper
    // ---------------------------------------------------------------------------

    /// <summary>
    /// Returns the Cell at <paramref name="cellRef"/> in <paramref name="worksheet"/>,
    /// creating Row and/or Cell nodes in sorted order if they do not already exist.
    ///
    /// OpenXML requires rows to be in ascending row-index order and cells within a
    /// row to be in ascending column order; violating this order corrupts the file.
    /// </summary>
    public static Cell EnsureCell(Worksheet worksheet, string cellRef)
    {
        var (col, rowIndex) = Parse(cellRef);
        string normalizedRef = ToA1(col, rowIndex);

        SheetData sheetData = worksheet.GetFirstChild<SheetData>()
            ?? throw new InvalidOperationException("Worksheet has no SheetData element");

        // Find or create the Row, maintaining ascending row index order.
        Row? row = FindRow(sheetData, (uint)rowIndex);
        if (row is null)
        {
            row = new Row { RowIndex = (uint)rowIndex };
            InsertRowInOrder(sheetData, row);
        }

        // Find or create the Cell, maintaining ascending column order within the row.
        Cell? cell = FindCell(row, normalizedRef);
        if (cell is null)
        {
            cell = new Cell { CellReference = normalizedRef };
            InsertCellInOrder(row, cell);
        }

        return cell;
    }

    // ---------------------------------------------------------------------------
    // Private helpers
    // ---------------------------------------------------------------------------

    private static Row? FindRow(SheetData sheetData, uint rowIndex)
    {
        foreach (var child in sheetData.Elements<Row>())
            if (child.RowIndex?.Value == rowIndex) return child;
        return null;
    }

    private static Cell? FindCell(Row row, string cellRef)
    {
        foreach (var child in row.Elements<Cell>())
            if (string.Equals(child.CellReference?.Value, cellRef, StringComparison.OrdinalIgnoreCase))
                return child;
        return null;
    }

    private static void InsertRowInOrder(SheetData sheetData, Row newRow)
    {
        uint newIndex = newRow.RowIndex?.Value ?? 0;

        // Find the first existing row with a higher index to insert before it.
        Row? refRow = null;
        foreach (var existingRow in sheetData.Elements<Row>())
        {
            if ((existingRow.RowIndex?.Value ?? 0) > newIndex)
            {
                refRow = existingRow;
                break;
            }
        }

        if (refRow is null)
            sheetData.AppendChild(newRow);
        else
            sheetData.InsertBefore(newRow, refRow);
    }

    private static void InsertCellInOrder(Row row, Cell newCell)
    {
        string newRef = newCell.CellReference?.Value ?? "";
        var (newCol, _) = Parse(newRef);

        // Find the first existing cell with a higher column index.
        Cell? refCell = null;
        foreach (var existing in row.Elements<Cell>())
        {
            string existRef = existing.CellReference?.Value ?? "";
            if (existRef.Length == 0) continue;
            try
            {
                var (existCol, _) = Parse(existRef);
                if (existCol > newCol)
                {
                    refCell = existing;
                    break;
                }
            }
            catch { /* skip malformed references */ }
        }

        if (refCell is null)
            row.AppendChild(newCell);
        else
            row.InsertBefore(newCell, refCell);
    }
}
