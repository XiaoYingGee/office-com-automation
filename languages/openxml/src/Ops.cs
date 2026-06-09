using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelOps;

public static class Ops
{
    // ---------------------------------------------------------------------------
    // Dispatch
    // ---------------------------------------------------------------------------

    public static object Dispatch(OpRequest req)
    {
        return req.Op switch
        {
            "cell.write"        => CellWrite(req),
            "range.write_bulk"  => RangeWriteBulk(req),
            "range.read"        => RangeRead(req),
            "range.clear"       => RangeClear(req),
            "range.merge"       => RangeMerge(req),
            "range.unmerge"     => RangeUnmerge(req),
            "range.copy_values" => RangeCopyValues(req),
            "inspect"           => Inspect(req),
            "set_format"        => SetFormat(req),
            "row.insert"        => RowInsert(req),
            "row.delete"        => RowDelete(req),
            "sheet.add"         => SheetAdd(req),
            "sheet.rename"      => SheetRename(req),
            "sheet.delete"      => SheetDelete(req),
            _ => throw new OpException(ErrorCategory.Unknown,
                $"unknown op: {req.Op}",
                hint: "supported: cell.write, range.read, range.write_bulk, range.clear, range.merge, range.unmerge, range.copy_values, inspect, set_format, row.insert, row.delete, sheet.add, sheet.rename, sheet.delete"),
        };
    }

    // ---------------------------------------------------------------------------
    // cell.write
    // ---------------------------------------------------------------------------

    private static object CellWrite(OpRequest req)
    {
        // Validate target.range
        string? rangeAddr = req.Target?.Range;
        if (string.IsNullOrWhiteSpace(rangeAddr))
            throw new OpException(ErrorCategory.InvalidArg, "cell.write requires target.range");

        // Parse params
        if (req.Params.ValueKind == JsonValueKind.Undefined || req.Params.ValueKind == JsonValueKind.Null)
            throw new OpException(ErrorCategory.InvalidArg, "cell.write requires params");

        if (!req.Params.TryGetProperty("kind", out var kindEl))
            throw new OpException(ErrorCategory.InvalidArg, "params.kind is required");
        string kind = kindEl.GetString()
            ?? throw new OpException(ErrorCategory.InvalidArg, $"params.kind must be a string (got {kindEl.ValueKind})");

        if (!req.Params.TryGetProperty("value", out var valueEl))
            throw new OpException(ErrorCategory.InvalidArg, "params.value is required");

        // Validate save_as format if present
        ValidateSaveAsFormat(req);

        using var doc = OpenOrCreate(req);

        var workbookPart = doc.WorkbookPart
            ?? throw new OpException(ErrorCategory.Unknown, "Workbook part missing");

        // Resolve worksheet
        WorksheetPart worksheetPart = ResolveWorksheetPart(workbookPart, req.Target?.Sheet);
        Worksheet worksheet = worksheetPart.Worksheet;

        // Get or create the target cell
        Cell cell = A1.EnsureCell(worksheet, rangeAddr);

        // Write value based on kind
        WriteCellValue(cell, kind, valueEl);

        worksheet.Save();

        return new { written = true };
    }

    // ---------------------------------------------------------------------------
    // range.write_bulk
    // ---------------------------------------------------------------------------

    private static object RangeWriteBulk(OpRequest req)
    {
        string? rangeAddr = req.Target?.Range;
        if (string.IsNullOrWhiteSpace(rangeAddr))
            throw new OpException(ErrorCategory.InvalidArg, "range.write_bulk requires target.range");

        // Require a range with ":"
        if (!rangeAddr.Contains(':'))
            throw new OpException(ErrorCategory.InvalidArg,
                $"range.write_bulk requires a multi-cell range (e.g. A1:B2), got: {rangeAddr}");

        // Parse params.values
        if (req.Params.ValueKind == JsonValueKind.Undefined || req.Params.ValueKind == JsonValueKind.Null)
            throw new OpException(ErrorCategory.InvalidArg, "range.write_bulk requires params");

        if (!req.Params.TryGetProperty("values", out var valuesEl) ||
            valuesEl.ValueKind != JsonValueKind.Array)
            throw new OpException(ErrorCategory.InvalidArg, "params.values must be a non-empty array");

        var rows = valuesEl.EnumerateArray().ToList();
        if (rows.Count == 0)
            throw new OpException(ErrorCategory.InvalidArg, "params.values must be a non-empty array");

        // Validate rectangular: all rows same length, non-empty
        int firstLen = -1;
        for (int r = 0; r < rows.Count; r++)
        {
            if (rows[r].ValueKind != JsonValueKind.Array)
                throw new OpException(ErrorCategory.InvalidArg,
                    $"params.values[{r}] must be an array");
            int len = rows[r].GetArrayLength();
            if (len == 0)
                throw new OpException(ErrorCategory.InvalidArg,
                    $"params.values[{r}] is an empty row; all rows must be non-empty");
            if (firstLen < 0) firstLen = len;
            else if (len != firstLen)
                throw new OpException(ErrorCategory.InvalidArg,
                    $"params.values[{r}] has {len} columns; expected {firstLen} (ragged array)");
        }

        int numRows = rows.Count;
        int numCols = firstLen;

        // Parse anchor from top-left of range
        string topLeft = rangeAddr.Split(':')[0];
        var (anchorCol, anchorRow) = A1.Parse(topLeft);

        ValidateSaveAsFormat(req);

        using var doc = OpenOrCreate(req);
        var workbookPart = doc.WorkbookPart
            ?? throw new OpException(ErrorCategory.Unknown, "Workbook part missing");
        WorksheetPart worksheetPart = ResolveWorksheetPart(workbookPart, req.Target?.Sheet);
        Worksheet worksheet = worksheetPart.Worksheet;

        for (int r = 0; r < numRows; r++)
        {
            var rowArr = rows[r].EnumerateArray().ToList();
            for (int c = 0; c < numCols; c++)
            {
                string cellRef = A1.ToA1(anchorCol + c, anchorRow + r);
                Cell cell = A1.EnsureCell(worksheet, cellRef);
                WriteJsonElementToCell(cell, rowArr[c]);
            }
        }

        worksheet.Save();
        return new { written = true, rows = numRows, cols = numCols };
    }

    // ---------------------------------------------------------------------------
    // range.clear
    // ---------------------------------------------------------------------------

    private static object RangeClear(OpRequest req)
    {
        // File must exist
        if (!File.Exists(req.Path))
            throw new OpException(ErrorCategory.FileNotFound,
                $"File not found: {req.Path}");

        string? rangeAddr = req.Target?.Range;
        if (string.IsNullOrWhiteSpace(rangeAddr))
            throw new OpException(ErrorCategory.InvalidArg, "range.clear requires target.range");

        // Determine mode (default: contents)
        string mode = "contents";
        if (req.Params.ValueKind != JsonValueKind.Undefined &&
            req.Params.ValueKind != JsonValueKind.Null &&
            req.Params.TryGetProperty("mode", out var modeEl) &&
            modeEl.ValueKind == JsonValueKind.String)
        {
            mode = modeEl.GetString()!;
        }

        if (mode != "contents" && mode != "all")
            throw new OpException(ErrorCategory.InvalidArg,
                $"params.mode must be 'contents' or 'all', got: {mode}");

        ValidateSaveAsFormat(req);

        using var doc = OpenDocument(req);
        var workbookPart = doc.WorkbookPart
            ?? throw new OpException(ErrorCategory.Unknown, "Workbook part missing");
        WorksheetPart worksheetPart = ResolveWorksheetPart(workbookPart, req.Target?.Sheet);
        Worksheet worksheet = worksheetPart.Worksheet;

        var sheetData = worksheet.GetFirstChild<SheetData>()
            ?? throw new InvalidOperationException("Worksheet has no SheetData element");

        // Enumerate cells to clear
        IEnumerable<string> cellRefs = EnumerateRangeCells(rangeAddr);

        foreach (string cellRef in cellRefs)
        {
            var (col, rowIdx) = A1.Parse(cellRef);
            string normRef = A1.ToA1(col, rowIdx);

            // Find existing cell (if any) and remove it
            foreach (var row in sheetData.Elements<Row>())
            {
                if (row.RowIndex?.Value != (uint)rowIdx) continue;

                Cell? found = null;
                foreach (var c in row.Elements<Cell>())
                {
                    if (string.Equals(c.CellReference?.Value, normRef, StringComparison.OrdinalIgnoreCase))
                    {
                        found = c;
                        break;
                    }
                }

                if (found is not null)
                {
                    // Remove the cell element entirely — cleanest way to clear both value and style
                    row.RemoveChild(found);
                }
                break;
            }
        }

        worksheet.Save();
        return new { cleared = true };
    }

    // ---------------------------------------------------------------------------
    // range.merge
    // ---------------------------------------------------------------------------

    private static object RangeMerge(OpRequest req)
    {
        if (!File.Exists(req.Path))
            throw new OpException(ErrorCategory.FileNotFound, $"File not found: {req.Path}");

        string? rangeAddr = req.Target?.Range;
        if (string.IsNullOrWhiteSpace(rangeAddr))
            throw new OpException(ErrorCategory.InvalidArg, "range.merge requires target.range");
        if (!rangeAddr.Contains(':'))
            throw new OpException(ErrorCategory.InvalidArg,
                "range.merge requires a multi-cell range (e.g. A1:C3)");

        ValidateSaveAsFormat(req);
        using var doc = OpenDocument(req);
        var workbookPart = doc.WorkbookPart
            ?? throw new OpException(ErrorCategory.Unknown, "Workbook part missing");
        WorksheetPart worksheetPart = ResolveWorksheetPart(workbookPart, req.Target?.Sheet);
        Worksheet worksheet = worksheetPart.Worksheet;

        var sheetData = worksheet.GetFirstChild<SheetData>()
            ?? throw new InvalidOperationException("Worksheet has no SheetData element");

        var mergeCells = worksheet.GetFirstChild<MergeCells>();
        if (mergeCells is null)
        {
            // Per the SpreadsheetML schema, mergeCells follows sheetData.
            mergeCells = new MergeCells();
            worksheet.InsertAfter(mergeCells, sheetData);
        }

        bool exists = mergeCells.Elements<MergeCell>()
            .Any(m => string.Equals(m.Reference?.Value, rangeAddr, StringComparison.OrdinalIgnoreCase));
        if (!exists)
            mergeCells.AppendChild(new MergeCell { Reference = rangeAddr });

        worksheet.Save();
        return new { merged = true };
    }

    // ---------------------------------------------------------------------------
    // range.unmerge
    // ---------------------------------------------------------------------------

    private static object RangeUnmerge(OpRequest req)
    {
        if (!File.Exists(req.Path))
            throw new OpException(ErrorCategory.FileNotFound, $"File not found: {req.Path}");

        string? rangeAddr = req.Target?.Range;
        if (string.IsNullOrWhiteSpace(rangeAddr))
            throw new OpException(ErrorCategory.InvalidArg, "range.unmerge requires target.range");

        ValidateSaveAsFormat(req);
        using var doc = OpenDocument(req);
        var workbookPart = doc.WorkbookPart
            ?? throw new OpException(ErrorCategory.Unknown, "Workbook part missing");
        WorksheetPart worksheetPart = ResolveWorksheetPart(workbookPart, req.Target?.Sheet);
        Worksheet worksheet = worksheetPart.Worksheet;

        var mergeCells = worksheet.GetFirstChild<MergeCells>();
        if (mergeCells is not null)
        {
            var toRemove = mergeCells.Elements<MergeCell>()
                .Where(m => string.Equals(m.Reference?.Value, rangeAddr, StringComparison.OrdinalIgnoreCase))
                .ToList();
            foreach (var m in toRemove) mergeCells.RemoveChild(m);
            if (!mergeCells.Elements<MergeCell>().Any())
                worksheet.RemoveChild(mergeCells);
        }

        worksheet.Save();
        return new { unmerged = true };
    }


    // ---------------------------------------------------------------------------
    // range.copy_values
    // ---------------------------------------------------------------------------

    private static object RangeCopyValues(OpRequest req)
    {
        // File must exist
        if (!File.Exists(req.Path))
            throw new OpException(ErrorCategory.FileNotFound,
                $"File not found: {req.Path}");

        string? sourceRef = req.Target?.Range;
        if (string.IsNullOrWhiteSpace(sourceRef))
            throw new OpException(ErrorCategory.InvalidArg, "range.copy_values requires target.range");

        // Get dest param
        if (req.Params.ValueKind == JsonValueKind.Undefined || req.Params.ValueKind == JsonValueKind.Null ||
            !req.Params.TryGetProperty("dest", out var destEl) ||
            destEl.ValueKind != JsonValueKind.String)
            throw new OpException(ErrorCategory.InvalidArg, "params.dest is required");

        string dest = destEl.GetString()!;
        if (string.IsNullOrWhiteSpace(dest))
            throw new OpException(ErrorCategory.InvalidArg, "params.dest must be a non-empty string");

        // No cross-sheet copies
        if (dest.Contains('!'))
            throw new OpException(ErrorCategory.InvalidArg,
                "params.dest must not contain a sheet reference (cross-sheet copy not supported)");

        ValidateSaveAsFormat(req);

        using var doc = OpenDocument(req);
        var workbookPart = doc.WorkbookPart
            ?? throw new OpException(ErrorCategory.Unknown, "Workbook part missing");
        WorksheetPart worksheetPart = ResolveWorksheetPart(workbookPart, req.Target?.Sheet);
        Worksheet worksheet = worksheetPart.Worksheet;

        // Read source cell value
        var sheetData = worksheet.GetFirstChild<SheetData>()
            ?? throw new InvalidOperationException("Worksheet has no SheetData element");

        var (srcCol, srcRow) = A1.Parse(sourceRef);
        string srcNorm = A1.ToA1(srcCol, srcRow);

        Cell? srcCell = null;
        foreach (var row in sheetData.Elements<Row>())
        {
            if (row.RowIndex?.Value != (uint)srcRow) continue;
            foreach (var c in row.Elements<Cell>())
            {
                if (string.Equals(c.CellReference?.Value, srcNorm, StringComparison.OrdinalIgnoreCase))
                {
                    srcCell = c;
                    break;
                }
            }
            break;
        }

        // Write to dest cell — preserving type
        Cell destCell = A1.EnsureCell(worksheet, dest);

        if (srcCell is null)
        {
            // Source is empty → dest becomes empty (remove value/formula)
            destCell.CellValue = null;
            destCell.DataType = null;
            destCell.RemoveAllChildren<InlineString>();
            destCell.RemoveAllChildren<CellFormula>();
        }
        else
        {
            CopyCell(srcCell, destCell, workbookPart);
        }

        worksheet.Save();
        return new { copied = true };
    }

    // ---------------------------------------------------------------------------
    // range.read
    // ---------------------------------------------------------------------------

    private static object RangeRead(OpRequest req)
    {
        if (!File.Exists(req.Path))
            throw new OpException(ErrorCategory.FileNotFound, $"File not found: {req.Path}");

        string? rangeAddr = req.Target?.Range;
        if (string.IsNullOrWhiteSpace(rangeAddr))
            throw new OpException(ErrorCategory.InvalidArg, "range.read requires target.range");

        string property = "value2";
        if (req.Params.ValueKind != JsonValueKind.Undefined &&
            req.Params.ValueKind != JsonValueKind.Null &&
            req.Params.TryGetProperty("property", out var propEl) &&
            propEl.ValueKind == JsonValueKind.String)
        {
            property = propEl.GetString()!;
        }

        using var doc = SpreadsheetDocument.Open(req.Path, isEditable: false);
        var workbookPart = doc.WorkbookPart
            ?? throw new OpException(ErrorCategory.Unknown, "Workbook part missing");
        WorksheetPart worksheetPart = ResolveWorksheetPart(workbookPart, req.Target?.Sheet);
        var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();

        var (col, row) = A1.Parse(rangeAddr);
        string normRef = A1.ToA1(col, row);

        Cell? cell = null;
        if (sheetData is not null)
        {
            foreach (var r in sheetData.Elements<Row>())
            {
                if (r.RowIndex?.Value != (uint)row) continue;
                foreach (var c in r.Elements<Cell>())
                {
                    if (string.Equals(c.CellReference?.Value, normRef, StringComparison.OrdinalIgnoreCase))
                    { cell = c; break; }
                }
                break;
            }
        }

        if (cell is null)
            return new { kind = "empty", value = (object?)null };

        if (property == "formula")
        {
            var formula = cell.GetFirstChild<CellFormula>();
            if (formula is not null)
                return new { kind = "string", value = (object?)("=" + formula.Text) };
        }

        return ReadCellValue(cell, workbookPart);
    }

    private static object ReadCellValue(Cell cell, WorkbookPart workbookPart)
    {
        var dataType = cell.DataType?.Value;

        if (dataType == CellValues.InlineString)
        {
            string? text = cell.GetFirstChild<InlineString>()?.GetFirstChild<Text>()?.Text;
            return new { kind = "string", value = (object?)(text ?? "") };
        }
        if (dataType == CellValues.SharedString)
        {
            string? idx = cell.CellValue?.Text;
            string resolved = "";
            if (idx is not null && int.TryParse(idx, out int i))
            {
                var sst = workbookPart.SharedStringTablePart?.SharedStringTable;
                resolved = sst?.Elements<SharedStringItem>().ElementAtOrDefault(i)?.InnerText ?? "";
            }
            return new { kind = "string", value = (object?)resolved };
        }
        if (dataType == CellValues.Boolean)
        {
            bool bv = cell.CellValue?.Text == "1";
            return new { kind = "bool", value = (object?)bv };
        }

        string? raw = cell.CellValue?.Text;
        if (raw is not null && double.TryParse(raw, System.Globalization.NumberStyles.Float,
            System.Globalization.CultureInfo.InvariantCulture, out double dv))
        {
            return new { kind = "number", value = (object?)dv };
        }

        if (raw is null)
            return new { kind = "empty", value = (object?)null };

        return new { kind = "string", value = (object?)raw };
    }

    // ---------------------------------------------------------------------------
    // inspect
    // ---------------------------------------------------------------------------

    private static object Inspect(OpRequest req)
    {
        if (!File.Exists(req.Path))
            throw new OpException(ErrorCategory.FileNotFound, $"File not found: {req.Path}");

        using var doc = SpreadsheetDocument.Open(req.Path, isEditable: false);
        var workbookPart = doc.WorkbookPart
            ?? throw new OpException(ErrorCategory.Unknown, "Workbook part missing");
        var sheets = workbookPart.Workbook.GetFirstChild<Sheets>();
        if (sheets is null)
            return new { sheets = Array.Empty<object>() };

        var result = new List<object>();
        int idx = 0;
        foreach (var s in sheets.Elements<Sheet>())
        {
            idx++;
            string name = s.Name?.Value ?? "";
            int rows = 0, cols = 0;
            string usedRange = "";

            string? relId = s.Id?.Value;
            if (relId is not null && workbookPart.GetPartById(relId) is WorksheetPart wsp)
            {
                var sheetData = wsp.Worksheet.GetFirstChild<SheetData>();
                if (sheetData is not null)
                {
                    int minRow = int.MaxValue, maxRow = 0, minCol = int.MaxValue, maxCol = 0;
                    foreach (var row in sheetData.Elements<Row>())
                    {
                        int ri = (int)(row.RowIndex?.Value ?? 0);
                        if (ri < minRow) minRow = ri;
                        if (ri > maxRow) maxRow = ri;
                        foreach (var cell in row.Elements<Cell>())
                        {
                            string? cr = cell.CellReference?.Value;
                            if (cr is null) continue;
                            try { var (c, _) = A1.Parse(cr); if (c < minCol) minCol = c; if (c > maxCol) maxCol = c; }
                            catch { }
                        }
                    }
                    if (maxRow > 0)
                    {
                        rows = maxRow - minRow + 1;
                        cols = maxCol - minCol + 1;
                        usedRange = $"{A1.ToA1(minCol, minRow)}:{A1.ToA1(maxCol, maxRow)}";
                    }
                }
            }

            result.Add(new { index = idx, name, used_range = usedRange, rows, cols });
        }

        return new { sheets = result };
    }

    // ---------------------------------------------------------------------------
    // set_format
    // ---------------------------------------------------------------------------

    private static object SetFormat(OpRequest req)
    {
        if (!File.Exists(req.Path))
            throw new OpException(ErrorCategory.FileNotFound, $"File not found: {req.Path}");

        string? rangeAddr = req.Target?.Range;
        if (string.IsNullOrWhiteSpace(rangeAddr))
            throw new OpException(ErrorCategory.InvalidArg, "set_format requires target.range");

        ValidateSaveAsFormat(req);
        using var doc = OpenDocument(req);
        var workbookPart = doc.WorkbookPart
            ?? throw new OpException(ErrorCategory.Unknown, "Workbook part missing");
        WorksheetPart worksheetPart = ResolveWorksheetPart(workbookPart, req.Target?.Sheet);
        Worksheet worksheet = worksheetPart.Worksheet;

        var stylesPart = workbookPart.WorkbookStylesPart;
        if (stylesPart is null)
        {
            stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
            stylesPart.Stylesheet = new Stylesheet(
                new Fonts(new Font()),
                new Fills(new Fill(new PatternFill { PatternType = PatternValues.None }),
                          new Fill(new PatternFill { PatternType = PatternValues.Gray125 })),
                new Borders(new Border()),
                new CellFormats(new CellFormat())
            );
            stylesPart.Stylesheet.Save();
        }
        var stylesheet = stylesPart.Stylesheet;

        bool? bold = null;
        double? fontSize = null;
        string? bgColorHex = null;

        if (req.Params.ValueKind != JsonValueKind.Undefined && req.Params.ValueKind != JsonValueKind.Null)
        {
            if (req.Params.TryGetProperty("bold", out var bEl) && (bEl.ValueKind == JsonValueKind.True || bEl.ValueKind == JsonValueKind.False))
                bold = bEl.GetBoolean();
            if (req.Params.TryGetProperty("font_size", out var fsEl) && fsEl.ValueKind == JsonValueKind.Number)
                fontSize = fsEl.GetDouble();
            if (req.Params.TryGetProperty("bg_color", out var bgEl) && bgEl.ValueKind == JsonValueKind.Number)
                bgColorHex = bgEl.GetInt64().ToString("X6");
        }

        var fonts = stylesheet.Fonts ?? stylesheet.AppendChild(new Fonts(new Font()));
        var fills = stylesheet.Fills ?? stylesheet.AppendChild(new Fills(
            new Fill(new PatternFill { PatternType = PatternValues.None }),
            new Fill(new PatternFill { PatternType = PatternValues.Gray125 })));
        var cellFormats = stylesheet.CellFormats ?? stylesheet.AppendChild(new CellFormats(new CellFormat()));

        var newFont = new Font();
        if (bold == true) newFont.AppendChild(new Bold());
        if (fontSize.HasValue) newFont.AppendChild(new FontSize { Val = fontSize.Value });
        fonts.AppendChild(newFont);
        uint fontId = (uint)(fonts.Count ?? 1) - 1;

        uint fillId = 0;
        if (bgColorHex is not null)
        {
            var newFill = new Fill(new PatternFill(new ForegroundColor { Rgb = new HexBinaryValue(bgColorHex) })
                { PatternType = PatternValues.Solid });
            fills.AppendChild(newFill);
            fillId = (uint)(fills.Count ?? 1) - 1;
        }

        var newCellFormat = new CellFormat { FontId = fontId, FillId = fillId, ApplyFont = true };
        if (fillId > 0) newCellFormat.ApplyFill = true;
        cellFormats.AppendChild(newCellFormat);
        uint styleIndex = (uint)(cellFormats.Count ?? 1) - 1;

        stylesheet.Save();

        foreach (string cellRef in EnumerateRangeCells(rangeAddr))
        {
            Cell cell = A1.EnsureCell(worksheet, cellRef);
            cell.StyleIndex = styleIndex;
        }

        worksheet.Save();
        return new { formatted = true };
    }

    // ---------------------------------------------------------------------------
    // row.insert
    // ---------------------------------------------------------------------------

    private static object RowInsert(OpRequest req)
    {
        if (!File.Exists(req.Path))
            throw new OpException(ErrorCategory.FileNotFound, $"File not found: {req.Path}");

        int row = 1, count = 1;
        if (req.Params.ValueKind != JsonValueKind.Undefined && req.Params.ValueKind != JsonValueKind.Null)
        {
            if (req.Params.TryGetProperty("row", out var rEl) && rEl.ValueKind == JsonValueKind.Number)
                row = rEl.GetInt32();
            if (req.Params.TryGetProperty("count", out var cEl) && cEl.ValueKind == JsonValueKind.Number)
                count = cEl.GetInt32();
        }

        ValidateSaveAsFormat(req);
        using var doc = OpenDocument(req);
        var workbookPart = doc.WorkbookPart
            ?? throw new OpException(ErrorCategory.Unknown, "Workbook part missing");
        WorksheetPart worksheetPart = ResolveWorksheetPart(workbookPart, req.Target?.Sheet);
        var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
        if (sheetData is null) { worksheetPart.Worksheet.Save(); return new { inserted = true, row, count }; }

        foreach (var r in sheetData.Elements<Row>().ToList())
        {
            uint ri = r.RowIndex?.Value ?? 0;
            if (ri >= (uint)row)
            {
                r.RowIndex = ri + (uint)count;
                foreach (var c in r.Elements<Cell>())
                {
                    string? cr = c.CellReference?.Value;
                    if (cr is null) continue;
                    var (col, _) = A1.Parse(cr);
                    c.CellReference = A1.ToA1(col, (int)(ri + (uint)count));
                }
            }
        }

        worksheetPart.Worksheet.Save();
        return new { inserted = true, row, count };
    }

    // ---------------------------------------------------------------------------
    // row.delete
    // ---------------------------------------------------------------------------

    private static object RowDelete(OpRequest req)
    {
        if (!File.Exists(req.Path))
            throw new OpException(ErrorCategory.FileNotFound, $"File not found: {req.Path}");

        int row = 1, count = 1;
        if (req.Params.ValueKind != JsonValueKind.Undefined && req.Params.ValueKind != JsonValueKind.Null)
        {
            if (req.Params.TryGetProperty("row", out var rEl) && rEl.ValueKind == JsonValueKind.Number)
                row = rEl.GetInt32();
            if (req.Params.TryGetProperty("count", out var cEl) && cEl.ValueKind == JsonValueKind.Number)
                count = cEl.GetInt32();
        }

        ValidateSaveAsFormat(req);
        using var doc = OpenDocument(req);
        var workbookPart = doc.WorkbookPart
            ?? throw new OpException(ErrorCategory.Unknown, "Workbook part missing");
        WorksheetPart worksheetPart = ResolveWorksheetPart(workbookPart, req.Target?.Sheet);
        var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
        if (sheetData is null) { worksheetPart.Worksheet.Save(); return new { deleted = true, row, count }; }

        var rowsToRemove = sheetData.Elements<Row>()
            .Where(r => r.RowIndex?.Value >= (uint)row && r.RowIndex?.Value < (uint)(row + count))
            .ToList();
        foreach (var r in rowsToRemove) sheetData.RemoveChild(r);

        foreach (var r in sheetData.Elements<Row>().ToList())
        {
            uint ri = r.RowIndex?.Value ?? 0;
            if (ri > (uint)row)
            {
                r.RowIndex = ri - (uint)count;
                foreach (var c in r.Elements<Cell>())
                {
                    string? cr = c.CellReference?.Value;
                    if (cr is null) continue;
                    var (col, _) = A1.Parse(cr);
                    c.CellReference = A1.ToA1(col, (int)(ri - (uint)count));
                }
            }
        }

        worksheetPart.Worksheet.Save();
        return new { deleted = true, row, count };
    }

    // ---------------------------------------------------------------------------
    // sheet.add
    // ---------------------------------------------------------------------------

    private static object SheetAdd(OpRequest req)
    {
        string? name = null;
        if (req.Params.ValueKind != JsonValueKind.Undefined && req.Params.ValueKind != JsonValueKind.Null &&
            req.Params.TryGetProperty("name", out var nEl) && nEl.ValueKind == JsonValueKind.String)
            name = nEl.GetString();

        ValidateSaveAsFormat(req);
        using var doc = OpenOrCreate(req);
        var workbookPart = doc.WorkbookPart
            ?? throw new OpException(ErrorCategory.Unknown, "Workbook part missing");

        var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        worksheetPart.Worksheet = new Worksheet(new SheetData());

        var sheets = workbookPart.Workbook.GetFirstChild<Sheets>()
            ?? workbookPart.Workbook.AppendChild(new Sheets());

        uint maxId = 0;
        foreach (var s in sheets.Elements<Sheet>())
            if (s.SheetId?.Value > maxId) maxId = s.SheetId.Value;

        string sheetName = name ?? $"Sheet{maxId + 2}";
        sheets.AppendChild(new Sheet
        {
            Id = workbookPart.GetIdOfPart(worksheetPart),
            SheetId = maxId + 1,
            Name = sheetName,
        });

        workbookPart.Workbook.Save();
        return new { added = true, name = sheetName };
    }

    // ---------------------------------------------------------------------------
    // sheet.rename
    // ---------------------------------------------------------------------------

    private static object SheetRename(OpRequest req)
    {
        if (!File.Exists(req.Path))
            throw new OpException(ErrorCategory.FileNotFound, $"File not found: {req.Path}");

        string? sheetName = req.Target?.Sheet;
        if (string.IsNullOrWhiteSpace(sheetName))
            throw new OpException(ErrorCategory.InvalidArg, "sheet.rename requires target.sheet");

        string? newName = null;
        if (req.Params.ValueKind != JsonValueKind.Undefined && req.Params.ValueKind != JsonValueKind.Null &&
            req.Params.TryGetProperty("name", out var nEl) && nEl.ValueKind == JsonValueKind.String)
            newName = nEl.GetString();
        if (string.IsNullOrWhiteSpace(newName))
            throw new OpException(ErrorCategory.InvalidArg, "sheet.rename requires params.name");

        ValidateSaveAsFormat(req);
        using var doc = OpenDocument(req);
        var workbookPart = doc.WorkbookPart
            ?? throw new OpException(ErrorCategory.Unknown, "Workbook part missing");
        var sheets = workbookPart.Workbook.GetFirstChild<Sheets>();
        if (sheets is null)
            throw new OpException(ErrorCategory.SheetNotFound, "No sheets in workbook");

        Sheet? target = null;
        foreach (var s in sheets.Elements<Sheet>())
            if (string.Equals(s.Name?.Value, sheetName, StringComparison.Ordinal))
            { target = s; break; }

        if (target is null)
            throw new OpException(ErrorCategory.SheetNotFound, $"Sheet not found: {sheetName}");

        target.Name = newName;
        workbookPart.Workbook.Save();
        return new { renamed = true, old_name = sheetName, new_name = newName };
    }

    // ---------------------------------------------------------------------------
    // sheet.delete
    // ---------------------------------------------------------------------------

    private static object SheetDelete(OpRequest req)
    {
        if (!File.Exists(req.Path))
            throw new OpException(ErrorCategory.FileNotFound, $"File not found: {req.Path}");

        string? sheetName = req.Target?.Sheet;
        if (string.IsNullOrWhiteSpace(sheetName))
            throw new OpException(ErrorCategory.InvalidArg, "sheet.delete requires target.sheet");

        ValidateSaveAsFormat(req);
        using var doc = OpenDocument(req);
        var workbookPart = doc.WorkbookPart
            ?? throw new OpException(ErrorCategory.Unknown, "Workbook part missing");
        var sheets = workbookPart.Workbook.GetFirstChild<Sheets>();
        if (sheets is null)
            throw new OpException(ErrorCategory.SheetNotFound, "No sheets in workbook");

        Sheet? target = null;
        foreach (var s in sheets.Elements<Sheet>())
            if (string.Equals(s.Name?.Value, sheetName, StringComparison.Ordinal))
            { target = s; break; }

        if (target is null)
            throw new OpException(ErrorCategory.SheetNotFound, $"Sheet not found: {sheetName}");

        string? relId = target.Id?.Value;
        sheets.RemoveChild(target);
        if (relId is not null)
            workbookPart.DeletePart(relId);

        workbookPart.Workbook.Save();
        return new { deleted = true, name = sheetName };
    }

    // ---------------------------------------------------------------------------
    // Helpers — open/save
    // ---------------------------------------------------------------------------

    private static void ValidateSaveAsFormat(OpRequest req)
    {
        if (req.SaveAs is not null)
        {
            var fmt = req.SaveAs.Format.ToLowerInvariant();
            if (fmt != "xlsx" && fmt != "xlsm")
                throw new OpException(ErrorCategory.UnsupportedFormat,
                    $"unsupported format: {req.SaveAs.Format}; OpenXML backend supports xlsx and xlsm only");
        }
    }

    private static SpreadsheetDocument OpenOrCreate(OpRequest req)
    {
        bool useSaveAs = req.SaveAs is not null &&
            !string.Equals(req.SaveAs.Path, req.Path, StringComparison.OrdinalIgnoreCase);
        string writePath = useSaveAs ? req.SaveAs!.Path : req.Path;

        if (File.Exists(req.Path))
        {
            if (useSaveAs)
            {
                string? dir = Path.GetDirectoryName(writePath);
                if (!string.IsNullOrEmpty(dir)) Directory.CreateDirectory(dir);
                File.Copy(req.Path, writePath, overwrite: true);
            }
            return SpreadsheetDocument.Open(writePath, isEditable: true);
        }
        else
        {
            string? dir = Path.GetDirectoryName(writePath);
            if (!string.IsNullOrEmpty(dir)) Directory.CreateDirectory(dir);
            return CreateNewWorkbook(writePath, req.Target?.Sheet ?? "Sheet1");
        }
    }

    private static SpreadsheetDocument OpenDocument(OpRequest req)
    {
        // For ops that require an existing file (clear, copy_values).
        // Caller has already checked File.Exists.
        bool useSaveAs = req.SaveAs is not null &&
            !string.Equals(req.SaveAs.Path, req.Path, StringComparison.OrdinalIgnoreCase);
        string writePath = useSaveAs ? req.SaveAs!.Path : req.Path;

        if (useSaveAs)
        {
            string? dir = Path.GetDirectoryName(writePath);
            if (!string.IsNullOrEmpty(dir)) Directory.CreateDirectory(dir);
            File.Copy(req.Path, writePath, overwrite: true);
        }

        return SpreadsheetDocument.Open(writePath, isEditable: true);
    }

    // ---------------------------------------------------------------------------
    // Helpers — cell value writing
    // ---------------------------------------------------------------------------

    private static void WriteCellValue(Cell cell, string kind, JsonElement valueEl)
    {
        switch (kind)
        {
            case "string":
            {
                string strVal = valueEl.ValueKind == JsonValueKind.String
                    ? valueEl.GetString()!
                    : valueEl.ToString();
                cell.DataType = CellValues.InlineString;
                cell.CellValue = null;
                cell.RemoveAllChildren<InlineString>();
                cell.RemoveAllChildren<CellFormula>();
                cell.AppendChild(new InlineString(new Text(strVal)));
                break;
            }
            case "number":
            {
                if (valueEl.ValueKind != JsonValueKind.Number)
                    throw new OpException(ErrorCategory.InvalidArg,
                        "params.value must be a JSON number for kind=number");
                double num = valueEl.GetDouble();
                cell.DataType = null;
                cell.RemoveAllChildren<InlineString>();
                cell.RemoveAllChildren<CellFormula>();
                cell.CellValue = new CellValue(num.ToString("R", System.Globalization.CultureInfo.InvariantCulture));
                break;
            }
            case "bool":
            {
                if (valueEl.ValueKind != JsonValueKind.True && valueEl.ValueKind != JsonValueKind.False)
                    throw new OpException(ErrorCategory.InvalidArg,
                        "params.value must be a JSON bool for kind=bool");
                bool boolVal = valueEl.GetBoolean();
                cell.DataType = CellValues.Boolean;
                cell.RemoveAllChildren<InlineString>();
                cell.RemoveAllChildren<CellFormula>();
                cell.CellValue = new CellValue(boolVal ? "1" : "0");
                break;
            }
            case "formula":
            {
                string rawFormula = valueEl.ValueKind == JsonValueKind.String
                    ? valueEl.GetString()!
                    : valueEl.ToString();
                // Strip leading "=" if present
                string formulaText = rawFormula.StartsWith('=')
                    ? rawFormula[1..]
                    : rawFormula;

                // Clear any prior value/type/inline-string
                cell.DataType = null;
                cell.CellValue = null;
                cell.RemoveAllChildren<InlineString>();
                cell.RemoveAllChildren<CellFormula>();
                // Set formula — no cached value (honest: Excel recalculates on open)
                cell.AppendChild(new CellFormula(formulaText));
                break;
            }
            default:
                throw new OpException(ErrorCategory.InvalidArg, $"unknown kind: {kind}");
        }
    }

    private static void WriteJsonElementToCell(Cell cell, JsonElement el)
    {
        switch (el.ValueKind)
        {
            case JsonValueKind.String:
                cell.DataType = CellValues.InlineString;
                cell.CellValue = null;
                cell.RemoveAllChildren<InlineString>();
                cell.RemoveAllChildren<CellFormula>();
                cell.AppendChild(new InlineString(new Text(el.GetString()!)));
                break;

            case JsonValueKind.Number:
                cell.DataType = null;
                cell.RemoveAllChildren<InlineString>();
                cell.RemoveAllChildren<CellFormula>();
                cell.CellValue = new CellValue(
                    el.GetDouble().ToString("R", System.Globalization.CultureInfo.InvariantCulture));
                break;

            case JsonValueKind.True:
            case JsonValueKind.False:
                cell.DataType = CellValues.Boolean;
                cell.RemoveAllChildren<InlineString>();
                cell.RemoveAllChildren<CellFormula>();
                cell.CellValue = new CellValue(el.ValueKind == JsonValueKind.True ? "1" : "0");
                break;

            case JsonValueKind.Null:
            case JsonValueKind.Undefined:
                // null → leave cell empty
                cell.DataType = null;
                cell.CellValue = null;
                cell.RemoveAllChildren<InlineString>();
                cell.RemoveAllChildren<CellFormula>();
                break;

            default:
                // Anything else (object, array) → store as string
                cell.DataType = CellValues.InlineString;
                cell.CellValue = null;
                cell.RemoveAllChildren<InlineString>();
                cell.RemoveAllChildren<CellFormula>();
                cell.AppendChild(new InlineString(new Text(el.ToString())));
                break;
        }
    }

    private static void CopyCell(Cell src, Cell dest, WorkbookPart workbookPart)
    {
        // Clear dest first
        dest.CellValue = null;
        dest.DataType = null;
        dest.RemoveAllChildren<InlineString>();
        dest.RemoveAllChildren<CellFormula>();

        var dataType = src.DataType?.Value;

        if (dataType == CellValues.InlineString)
        {
            string? text = src.GetFirstChild<InlineString>()?.GetFirstChild<Text>()?.Text;
            dest.DataType = CellValues.InlineString;
            dest.AppendChild(new InlineString(new Text(text ?? "")));
        }
        else if (dataType == CellValues.SharedString)
        {
            // Resolve shared string and write as inline string
            string? idx = src.CellValue?.Text;
            string resolved = "";
            if (idx is not null && int.TryParse(idx, out int i))
            {
                var sst = workbookPart.SharedStringTablePart?.SharedStringTable;
                resolved = sst?.Elements<SharedStringItem>().ElementAtOrDefault(i)?.InnerText ?? "";
            }
            dest.DataType = CellValues.InlineString;
            dest.AppendChild(new InlineString(new Text(resolved)));
        }
        else if (dataType == CellValues.Boolean)
        {
            dest.DataType = CellValues.Boolean;
            dest.CellValue = new CellValue(src.CellValue?.Text ?? "0");
        }
        else
        {
            // Number, formula result, or empty — copy CellValue text as-is (numeric)
            string? raw = src.CellValue?.Text;
            if (raw is not null)
            {
                dest.DataType = null; // numeric
                dest.CellValue = new CellValue(raw);
            }
            // else: source was empty, dest stays empty
        }
    }

    // ---------------------------------------------------------------------------
    // Helpers — range enumeration
    // ---------------------------------------------------------------------------

    private static IEnumerable<string> EnumerateRangeCells(string rangeAddr)
    {
        if (!rangeAddr.Contains(':'))
        {
            // Single cell
            yield return rangeAddr;
            yield break;
        }

        string[] parts = rangeAddr.Split(':');
        var (col1, row1) = A1.Parse(parts[0]);
        var (col2, row2) = A1.Parse(parts[1]);

        int minCol = Math.Min(col1, col2);
        int maxCol = Math.Max(col1, col2);
        int minRow = Math.Min(row1, row2);
        int maxRow = Math.Max(row1, row2);

        for (int r = minRow; r <= maxRow; r++)
            for (int c = minCol; c <= maxCol; c++)
                yield return A1.ToA1(c, r);
    }

    // ---------------------------------------------------------------------------
    // Helpers — workbook creation / sheet resolution
    // ---------------------------------------------------------------------------

    private static SpreadsheetDocument CreateNewWorkbook(string path, string sheetName)
    {
        var doc = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook);

        var workbookPart = doc.AddWorkbookPart();
        workbookPart.Workbook = new Workbook();

        var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        worksheetPart.Worksheet = new Worksheet(new SheetData());

        var sheets = workbookPart.Workbook.AppendChild(new Sheets());
        sheets.AppendChild(new Sheet
        {
            Id = workbookPart.GetIdOfPart(worksheetPart),
            SheetId = 1,
            Name = sheetName,
        });

        workbookPart.Workbook.Save();

        return doc;
    }

    private static WorksheetPart ResolveWorksheetPart(WorkbookPart workbookPart, string? sheetName)
    {
        var sheets = workbookPart.Workbook.GetFirstChild<Sheets>();
        if (sheets is null)
            throw new OpException(ErrorCategory.Unknown, "Workbook has no Sheets element");

        Sheet? targetSheet = null;

        if (!string.IsNullOrWhiteSpace(sheetName))
        {
            // Try exact match first
            foreach (var s in sheets.Elements<Sheet>())
            {
                if (string.Equals(s.Name?.Value, sheetName, StringComparison.Ordinal))
                {
                    targetSheet = s;
                    break;
                }
            }
            // Fall back to first sheet — INTENTIONAL: mirrors Rust resolve_sheet_write
            if (targetSheet is null)
                targetSheet = sheets.Elements<Sheet>().FirstOrDefault();
        }
        else
        {
            targetSheet = sheets.Elements<Sheet>().FirstOrDefault();
        }

        if (targetSheet is null)
            throw new OpException(ErrorCategory.SheetNotFound, "No sheets found in workbook");

        string? relId = targetSheet.Id?.Value;
        if (string.IsNullOrEmpty(relId))
            throw new OpException(ErrorCategory.Unknown, "Sheet relationship ID is null");

        return workbookPart.GetPartById(relId) as WorksheetPart
            ?? throw new OpException(ErrorCategory.Unknown,
                $"Could not resolve WorksheetPart for sheet '{targetSheet.Name?.Value}'");
    }
}
