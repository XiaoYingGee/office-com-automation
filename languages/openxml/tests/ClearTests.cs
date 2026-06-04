using System.Diagnostics;
using System.Text.Json;
using Xunit;

namespace OpenXmlOps.Tests;

public class ClearTests
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
        Path.Combine(Path.GetTempPath(), $"ExcelOpsTest_Clear_{name}_{Guid.NewGuid():N}.xlsx");

    [Fact]
    public void Clear_contents_removesValue()
    {
        string path = TempXlsx("contents");
        try
        {
            // First write a value
            string writeReq = JsonSerializer.Serialize(new
            {
                op = "cell.write",
                path,
                target = new { sheet = "Sheet1", range = "A1" },
                @params = new { kind = "string", value = "to-be-cleared" },
            });
            Assert.Contains("\"ok\":true", RunOp(writeReq));
            Assert.Equal("to-be-cleared", OpenXmlReadback.ReadCellString(path, "Sheet1", "A1"));

            // Now clear
            string clearReq = JsonSerializer.Serialize(new
            {
                op = "range.clear",
                path,
                target = new { sheet = "Sheet1", range = "A1" },
                @params = new { mode = "contents" },
            });
            string raw = RunOp(clearReq);
            Assert.Contains("\"ok\":true", raw);
            Assert.Contains("\"cleared\":true", raw);

            // Readback should be null/empty
            string? val = OpenXmlReadback.ReadCellString(path, "Sheet1", "A1");
            Assert.True(val is null || val == "", $"Expected empty but got: {val}");
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public void Clear_all_removesCell()
    {
        string path = TempXlsx("all");
        try
        {
            string writeReq = JsonSerializer.Serialize(new
            {
                op = "cell.write",
                path,
                target = new { sheet = "Sheet1", range = "B2" },
                @params = new { kind = "number", value = 99 },
            });
            Assert.Contains("\"ok\":true", RunOp(writeReq));

            string clearReq = JsonSerializer.Serialize(new
            {
                op = "range.clear",
                path,
                target = new { sheet = "Sheet1", range = "B2" },
                @params = new { mode = "all" },
            });
            string raw = RunOp(clearReq);
            Assert.Contains("\"ok\":true", raw);

            string? val = OpenXmlReadback.ReadCellString(path, "Sheet1", "B2");
            Assert.True(val is null || val == "", $"Expected empty but got: {val}");
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public void Clear_defaultMode_contents()
    {
        string path = TempXlsx("default");
        try
        {
            string writeReq = JsonSerializer.Serialize(new
            {
                op = "cell.write",
                path,
                target = new { sheet = "Sheet1", range = "C3" },
                @params = new { kind = "string", value = "gone" },
            });
            Assert.Contains("\"ok\":true", RunOp(writeReq));

            // No mode param → defaults to contents
            string clearReq = JsonSerializer.Serialize(new
            {
                op = "range.clear",
                path,
                target = new { sheet = "Sheet1", range = "C3" },
                @params = new { },
            });
            string raw = RunOp(clearReq);
            Assert.Contains("\"ok\":true", raw);

            string? val = OpenXmlReadback.ReadCellString(path, "Sheet1", "C3");
            Assert.True(val is null || val == "", $"Expected empty but got: {val}");
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public void Clear_missingFile_returns_FileNotFound()
    {
        string path = TempXlsx("missing");
        // Do NOT create the file
        string clearReq = JsonSerializer.Serialize(new
        {
            op = "range.clear",
            path,
            target = new { sheet = "Sheet1", range = "A1" },
            @params = new { mode = "contents" },
        });
        string raw = RunOp(clearReq);
        Assert.Contains("\"ok\":false", raw);
        Assert.Contains("FileNotFound", raw);
    }

    [Fact]
    public void Clear_range_multiCell_clearsAll()
    {
        string path = TempXlsx("multi");
        try
        {
            // Write A1 and B1
            foreach (var (cell, val) in new[] { ("A1", "x"), ("B1", "y") })
            {
                string wr = JsonSerializer.Serialize(new
                {
                    op = "cell.write",
                    path,
                    target = new { sheet = "Sheet1", range = cell },
                    @params = new { kind = "string", value = val },
                });
                Assert.Contains("\"ok\":true", RunOp(wr));
            }

            // Clear range A1:B1
            string clearReq = JsonSerializer.Serialize(new
            {
                op = "range.clear",
                path,
                target = new { sheet = "Sheet1", range = "A1:B1" },
                @params = new { mode = "contents" },
            });
            Assert.Contains("\"ok\":true", RunOp(clearReq));

            string? v1 = OpenXmlReadback.ReadCellString(path, "Sheet1", "A1");
            string? v2 = OpenXmlReadback.ReadCellString(path, "Sheet1", "B1");
            Assert.True(v1 is null || v1 == "", $"A1 expected empty but got: {v1}");
            Assert.True(v2 is null || v2 == "", $"B1 expected empty but got: {v2}");
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }
}
