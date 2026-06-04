using System.Diagnostics;
using System.Text.Json;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Xunit;

namespace OpenXmlOps.Tests;

public class FormulaTests
{
    private static string FindExe() => CellWriteTests.FindExeHelper();

    private static string RunOp(string json)
    {
        string exe = FindExe();
        var psi = new ProcessStartInfo(exe)
        {
            RedirectStandardInput = true,
            RedirectStandardOutput = true,
            RedirectStandardError = false,
            UseShellExecute = false,
        };
        using var proc = Process.Start(psi)
            ?? throw new InvalidOperationException("Failed to start ExcelOps.exe");
        proc.StandardInput.Write(json);
        proc.StandardInput.Close();
        string output = proc.StandardOutput.ReadToEnd();
        proc.WaitForExit();
        return output;
    }

    private static string TempXlsx(string name) =>
        Path.Combine(Path.GetTempPath(), $"ExcelOpsTest_Formula_{name}_{Guid.NewGuid():N}.xlsx");

    [Fact]
    public void WriteFormula_cellHasF_element_and_no_V_element()
    {
        string path = TempXlsx("formula");
        try
        {
            string req = JsonSerializer.Serialize(new
            {
                op = "cell.write",
                path,
                target = new { sheet = "Sheet1", range = "A1" },
                @params = new { kind = "formula", value = "=2+3" },
            });

            string raw = RunOp(req);
            Assert.Contains("\"ok\":true", raw);

            // Open and inspect raw XML
            using var doc = SpreadsheetDocument.Open(path, isEditable: false);
            var wbPart = doc.WorkbookPart!;
            var sheet = wbPart.Workbook.GetFirstChild<Sheets>()!.Elements<Sheet>().First();
            var wsPart = (WorksheetPart)wbPart.GetPartById(sheet.Id!.Value!);
            var sheetData = wsPart.Worksheet.GetFirstChild<SheetData>()!;

            Cell? cell = null;
            foreach (var row in sheetData.Elements<Row>())
                foreach (var c in row.Elements<Cell>())
                    if (string.Equals(c.CellReference?.Value, "A1", StringComparison.OrdinalIgnoreCase))
                    { cell = c; break; }

            Assert.NotNull(cell);

            // Must have a CellFormula with text "2+3" (leading = stripped)
            var formula = cell!.GetFirstChild<CellFormula>();
            Assert.NotNull(formula);
            Assert.Equal("2+3", formula!.Text);

            // Must NOT have a CellValue (no stale cached result)
            Assert.Null(cell.CellValue);
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public void WriteFormula_noLeadingEquals_alsoWorks()
    {
        string path = TempXlsx("formula_noeq");
        try
        {
            string req = JsonSerializer.Serialize(new
            {
                op = "cell.write",
                path,
                target = new { sheet = "Sheet1", range = "B2" },
                @params = new { kind = "formula", value = "SUM(A1:A5)" },
            });

            string raw = RunOp(req);
            Assert.Contains("\"ok\":true", raw);

            using var doc = SpreadsheetDocument.Open(path, isEditable: false);
            var wbPart = doc.WorkbookPart!;
            var sheet = wbPart.Workbook.GetFirstChild<Sheets>()!.Elements<Sheet>().First();
            var wsPart = (WorksheetPart)wbPart.GetPartById(sheet.Id!.Value!);
            var sheetData = wsPart.Worksheet.GetFirstChild<SheetData>()!;

            Cell? cell = null;
            foreach (var row in sheetData.Elements<Row>())
                foreach (var c in row.Elements<Cell>())
                    if (string.Equals(c.CellReference?.Value, "B2", StringComparison.OrdinalIgnoreCase))
                    { cell = c; break; }

            Assert.NotNull(cell);
            var formula = cell!.GetFirstChild<CellFormula>();
            Assert.NotNull(formula);
            Assert.Equal("SUM(A1:A5)", formula!.Text);
            Assert.Null(cell.CellValue);
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }
}
