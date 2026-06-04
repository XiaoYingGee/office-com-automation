using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OpenXmlOps.Tests;

/// <summary>
/// Non-Excel oracle: opens an xlsx file with the OpenXML SDK and reads back
/// the raw stored value of a cell without invoking Excel.
/// </summary>
public static class OpenXmlReadback
{
    /// <summary>
    /// Open <paramref name="path"/>, find <paramref name="sheetName"/>,
    /// locate <paramref name="cellRef"/> and return its displayed/stored value as a string.
    ///
    /// - Numbers: returned as the raw CellValue text (e.g. "42.5").
    /// - Booleans: returned as "1" or "0" (OpenXML stores TRUE as 1, FALSE as 0).
    /// - Inline strings (t="inlineStr"): returned from the &lt;is&gt;&lt;t&gt; element.
    /// - Shared strings (t="s"): resolved via SharedStringTable.
    /// - Empty / missing cells: returns null.
    /// </summary>
    public static string? ReadCellString(string path, string sheetName, string cellRef)
    {
        using var doc = SpreadsheetDocument.Open(path, isEditable: false);

        var workbookPart = doc.WorkbookPart
            ?? throw new InvalidOperationException("No WorkbookPart");

        // Resolve sheet by name
        var sheets = workbookPart.Workbook.GetFirstChild<Sheets>();
        if (sheets is null) return null;

        Sheet? targetSheet = null;
        foreach (var s in sheets.Elements<Sheet>())
        {
            if (string.Equals(s.Name?.Value, sheetName, StringComparison.Ordinal))
            {
                targetSheet = s;
                break;
            }
        }
        if (targetSheet is null) return null;

        string? relId = targetSheet.Id?.Value;
        if (string.IsNullOrEmpty(relId)) return null;

        var worksheetPart = workbookPart.GetPartById(relId) as WorksheetPart;
        if (worksheetPart is null) return null;

        var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
        if (sheetData is null) return null;

        // Normalise cellRef so lookup is case-insensitive
        string normRef = cellRef.ToUpperInvariant();

        Cell? cell = null;
        foreach (var row in sheetData.Elements<Row>())
        {
            foreach (var c in row.Elements<Cell>())
            {
                if (string.Equals(c.CellReference?.Value, normRef, StringComparison.OrdinalIgnoreCase))
                {
                    cell = c;
                    break;
                }
            }
            if (cell is not null) break;
        }

        if (cell is null) return null;

        // Resolve value based on DataType
        var dataType = cell.DataType?.Value;

        if (dataType == CellValues.InlineString)
        {
            return cell.GetFirstChild<InlineString>()?.GetFirstChild<Text>()?.Text;
        }

        if (dataType == CellValues.SharedString)
        {
            string? idx = cell.CellValue?.Text;
            if (idx is null) return null;
            if (!int.TryParse(idx, out int i)) return null;
            var sst = workbookPart.SharedStringTablePart?.SharedStringTable;
            if (sst is null) return null;
            return sst.Elements<SharedStringItem>().ElementAtOrDefault(i)?.InnerText;
        }

        // Boolean, number, formula result — return raw CellValue text
        return cell.CellValue?.Text;
    }
}
