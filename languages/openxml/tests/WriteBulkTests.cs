using System.Diagnostics;
using System.Text.Json;
using Xunit;

namespace OpenXmlOps.Tests;

public class WriteBulkTests
{
    private static string RunOp(string json)
    {
        string exe = CellWriteTests.FindExeHelper();
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
        Path.Combine(Path.GetTempPath(), $"ExcelOpsTest_Bulk_{name}_{Guid.NewGuid():N}.xlsx");

    [Fact]
    public void WriteBulk_2x2_readback_correct()
    {
        string path = TempXlsx("basic");
        try
        {
            // [[10,20],[30,40]] anchored at A5:B6
            string req = JsonSerializer.Serialize(new
            {
                op = "range.write_bulk",
                path,
                target = new { sheet = "Sheet1", range = "A5:B6" },
                @params = new { values = new[] { new object[] { 10, 20 }, new object[] { 30, 40 } } },
            });

            string raw = RunOp(req);
            Assert.Contains("\"ok\":true", raw);
            Assert.Contains("\"written\":true", raw);
            Assert.Contains("\"rows\":2", raw);
            Assert.Contains("\"cols\":2", raw);

            Assert.Equal("10", OpenXmlReadback.ReadCellString(path, "Sheet1", "A5"));
            Assert.Equal("20", OpenXmlReadback.ReadCellString(path, "Sheet1", "B5"));
            Assert.Equal("30", OpenXmlReadback.ReadCellString(path, "Sheet1", "A6"));
            Assert.Equal("40", OpenXmlReadback.ReadCellString(path, "Sheet1", "B6"));
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public void WriteBulk_stringAndBool_readback()
    {
        string path = TempXlsx("types");
        try
        {
            string req = JsonSerializer.Serialize(new
            {
                op = "range.write_bulk",
                path,
                target = new { sheet = "Sheet1", range = "A1:C1" },
                @params = new { values = new[] { new object[] { "hello", true, 3.14 } } },
            });

            string raw = RunOp(req);
            Assert.Contains("\"ok\":true", raw);

            Assert.Equal("hello", OpenXmlReadback.ReadCellString(path, "Sheet1", "A1"));
            Assert.Equal("1", OpenXmlReadback.ReadCellString(path, "Sheet1", "B1"));
            Assert.Equal("3.14", OpenXmlReadback.ReadCellString(path, "Sheet1", "C1"));
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public void WriteBulk_raggedArray_returns_InvalidArg()
    {
        string path = TempXlsx("ragged");
        try
        {
            // Row 1 has 2 cols, Row 2 has 3 cols — ragged
            string req = "{\"op\":\"range.write_bulk\",\"path\":" + JsonSerializer.Serialize(path) +
                         ",\"target\":{\"sheet\":\"Sheet1\",\"range\":\"A1:C2\"},\"params\":{\"values\":[[1,2],[3,4,5]]}}";


            string raw = RunOp(req);
            Assert.Contains("\"ok\":false", raw);
            Assert.Contains("InvalidArg", raw);
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public void WriteBulk_singleCellRange_returns_InvalidArg()
    {
        string path = TempXlsx("single");
        try
        {
            string req = JsonSerializer.Serialize(new
            {
                op = "range.write_bulk",
                path,
                target = new { sheet = "Sheet1", range = "A1" }, // no colon
                @params = new { values = new[] { new object[] { 1 } } },
            });

            string raw = RunOp(req);
            Assert.Contains("\"ok\":false", raw);
            Assert.Contains("InvalidArg", raw);
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public void WriteBulk_emptyValues_returns_InvalidArg()
    {
        string path = TempXlsx("empty");
        try
        {
            string req = JsonSerializer.Serialize(new
            {
                op = "range.write_bulk",
                path,
                target = new { sheet = "Sheet1", range = "A1:B2" },
                @params = new { values = Array.Empty<object[]>() },
            });

            string raw = RunOp(req);
            Assert.Contains("\"ok\":false", raw);
            Assert.Contains("InvalidArg", raw);
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }
}
