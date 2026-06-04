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
            "cell.write" => CellWrite(req),
            _ => throw new OpException(ErrorCategory.Unknown,
                $"unknown op: {req.Op}",
                hint: "supported: cell.write"),
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
            ?? throw new OpException(ErrorCategory.InvalidArg, "params.kind must be a string");

        if (!req.Params.TryGetProperty("value", out var valueEl))
            throw new OpException(ErrorCategory.InvalidArg, "params.value is required");

        // Validate save_as format if present
        if (req.SaveAs is not null)
        {
            var fmt = req.SaveAs.Format.ToLowerInvariant();
            if (fmt != "xlsx" && fmt != "xlsm")
                throw new OpException(ErrorCategory.UnsupportedFormat,
                    $"unsupported format: {req.SaveAs.Format}; OpenXML backend supports xlsx and xlsm only");
        }

        // Determine output path
        string outputPath = req.SaveAs?.Path ?? req.Path;
        bool useSaveAs = req.SaveAs is not null &&
            !string.Equals(req.SaveAs.Path, req.Path, StringComparison.OrdinalIgnoreCase);

        // Write path: we always write to a temp-or-final path then optionally copy
        // For simplicity: write to outputPath directly (covers both cases since save_as.path may differ).
        string writePath = useSaveAs ? req.SaveAs!.Path : req.Path;

        SpreadsheetDocument doc;
        if (File.Exists(req.Path))
        {
            // If save_as differs from source, copy source to dest first so we edit the copy
            if (useSaveAs)
            {
                string? dir = Path.GetDirectoryName(writePath);
                if (!string.IsNullOrEmpty(dir)) Directory.CreateDirectory(dir);
                File.Copy(req.Path, writePath, overwrite: true);
            }
            doc = SpreadsheetDocument.Open(writePath, isEditable: true);
        }
        else
        {
            // Create new workbook
            string? dir = Path.GetDirectoryName(writePath);
            if (!string.IsNullOrEmpty(dir)) Directory.CreateDirectory(dir);
            doc = CreateNewWorkbook(writePath, req.Target?.Sheet ?? "Sheet1");
        }

        using (doc)
        {
            var workbookPart = doc.WorkbookPart
                ?? throw new OpException(ErrorCategory.Unknown, "Workbook part missing");

            // Resolve worksheet
            WorksheetPart worksheetPart = ResolveWorksheetPart(workbookPart, req.Target?.Sheet);
            Worksheet worksheet = worksheetPart.Worksheet;

            // Get or create the target cell
            Cell cell = A1.EnsureCell(worksheet, rangeAddr);

            // Write value based on kind
            switch (kind)
            {
                case "string":
                {
                    string strVal = valueEl.ValueKind == JsonValueKind.String
                        ? valueEl.GetString()!
                        : valueEl.ToString();
                    // Use inline string (t="inlineStr") — avoids shared-string-table complexity
                    cell.DataType = CellValues.InlineString;
                    cell.CellValue = null;
                    // Remove any existing InlineString child, then add fresh one
                    cell.RemoveAllChildren<InlineString>();
                    cell.AppendChild(new InlineString(new Text(strVal)));
                    break;
                }
                case "number":
                {
                    if (valueEl.ValueKind != JsonValueKind.Number)
                        throw new OpException(ErrorCategory.InvalidArg,
                            "params.value must be a JSON number for kind=number");
                    double num = valueEl.GetDouble();
                    cell.DataType = null; // numeric cells have no DataType attribute
                    cell.RemoveAllChildren<InlineString>();
                    // Use invariant culture to avoid locale-dependent decimal separators
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
                    cell.CellValue = new CellValue(boolVal ? "1" : "0");
                    break;
                }
                case "formula":
                    throw new OpException(ErrorCategory.UnsupportedFormat,
                        "kind=formula is not supported in Chunk 1; implement in Chunk 2");
                default:
                    throw new OpException(ErrorCategory.InvalidArg, $"unknown kind: {kind}");
            }

            worksheet.Save();
        }

        return new { written = true };
    }

    // ---------------------------------------------------------------------------
    // Helpers
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
            // Fall back to first sheet (mirrors Rust resolve_sheet_write behavior)
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
            ?? throw new OpException(ErrorCategory.Unknown, $"Could not resolve WorksheetPart for sheet '{targetSheet.Name?.Value}'");
    }
}
