using System.Diagnostics;
using System.Text.Json;
using Xunit;

namespace OpenXmlOps.Tests;

public class CopyValuesTests
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
        Path.Combine(Path.GetTempPath(), $"ExcelOpsTest_Copy_{name}_{Guid.NewGuid():N}.xlsx");

    [Fact]
    public void CopyValues_number_H1_to_H2()
    {
        string path = TempXlsx("number");
        try
        {
            // Write H1 = 7
            string writeReq = JsonSerializer.Serialize(new
            {
                op = "cell.write",
                path,
                target = new { sheet = "Sheet1", range = "H1" },
                @params = new { kind = "number", value = 7 },
            });
            Assert.Contains("\"ok\":true", RunOp(writeReq));

            // Copy H1 → H2
            string copyReq = JsonSerializer.Serialize(new
            {
                op = "range.copy_values",
                path,
                target = new { sheet = "Sheet1", range = "H1" },
                @params = new { dest = "H2" },
            });
            string raw = RunOp(copyReq);
            Assert.Contains("\"ok\":true", raw);
            Assert.Contains("\"copied\":true", raw);

            // H2 should now read as "7"
            string? val = OpenXmlReadback.ReadCellString(path, "Sheet1", "H2");
            Assert.Equal("7", val);
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public void CopyValues_string_preserved()
    {
        string path = TempXlsx("string");
        try
        {
            string writeReq = JsonSerializer.Serialize(new
            {
                op = "cell.write",
                path,
                target = new { sheet = "Sheet1", range = "A1" },
                @params = new { kind = "string", value = "copied-string" },
            });
            Assert.Contains("\"ok\":true", RunOp(writeReq));

            string copyReq = JsonSerializer.Serialize(new
            {
                op = "range.copy_values",
                path,
                target = new { sheet = "Sheet1", range = "A1" },
                @params = new { dest = "A2" },
            });
            Assert.Contains("\"ok\":true", RunOp(copyReq));

            Assert.Equal("copied-string", OpenXmlReadback.ReadCellString(path, "Sheet1", "A2"));
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public void CopyValues_bool_preserved()
    {
        string path = TempXlsx("bool");
        try
        {
            string writeReq = JsonSerializer.Serialize(new
            {
                op = "cell.write",
                path,
                target = new { sheet = "Sheet1", range = "A1" },
                @params = new { kind = "bool", value = true },
            });
            Assert.Contains("\"ok\":true", RunOp(writeReq));

            string copyReq = JsonSerializer.Serialize(new
            {
                op = "range.copy_values",
                path,
                target = new { sheet = "Sheet1", range = "A1" },
                @params = new { dest = "A3" },
            });
            Assert.Contains("\"ok\":true", RunOp(copyReq));

            Assert.Equal("1", OpenXmlReadback.ReadCellString(path, "Sheet1", "A3"));
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public void CopyValues_missingDest_returns_InvalidArg()
    {
        string path = TempXlsx("nodest");
        try
        {
            // Create file first
            string writeReq = JsonSerializer.Serialize(new
            {
                op = "cell.write",
                path,
                target = new { sheet = "Sheet1", range = "A1" },
                @params = new { kind = "number", value = 1 },
            });
            Assert.Contains("\"ok\":true", RunOp(writeReq));

            // Copy without dest
            string copyReq = JsonSerializer.Serialize(new
            {
                op = "range.copy_values",
                path,
                target = new { sheet = "Sheet1", range = "A1" },
                @params = new { },
            });
            string raw = RunOp(copyReq);
            Assert.Contains("\"ok\":false", raw);
            Assert.Contains("InvalidArg", raw);
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public void CopyValues_missingFile_returns_FileNotFound()
    {
        string path = TempXlsx("nofile");
        // File does NOT exist
        string copyReq = JsonSerializer.Serialize(new
        {
            op = "range.copy_values",
            path,
            target = new { sheet = "Sheet1", range = "A1" },
            @params = new { dest = "A2" },
        });
        string raw = RunOp(copyReq);
        Assert.Contains("\"ok\":false", raw);
        Assert.Contains("FileNotFound", raw);
    }

    [Fact]
    public void CopyValues_destWithSheetRef_returns_InvalidArg()
    {
        string path = TempXlsx("sheetref");
        try
        {
            string writeReq = JsonSerializer.Serialize(new
            {
                op = "cell.write",
                path,
                target = new { sheet = "Sheet1", range = "A1" },
                @params = new { kind = "number", value = 5 },
            });
            Assert.Contains("\"ok\":true", RunOp(writeReq));

            // dest with sheet reference → InvalidArg
            string copyReq = JsonSerializer.Serialize(new
            {
                op = "range.copy_values",
                path,
                target = new { sheet = "Sheet1", range = "A1" },
                @params = new { dest = "Sheet2!A1" },
            });
            string raw = RunOp(copyReq);
            Assert.Contains("\"ok\":false", raw);
            Assert.Contains("InvalidArg", raw);
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }
}
