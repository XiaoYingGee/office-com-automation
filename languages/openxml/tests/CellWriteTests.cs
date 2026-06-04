using System.Diagnostics;
using System.Text.Json;
using Xunit;

namespace OpenXmlOps.Tests;

public class CellWriteTests
{
    // ---------------------------------------------------------------------------
    // Subprocess helper
    // ---------------------------------------------------------------------------

    /// <summary>
    /// Locate ExcelOps.exe.  The test project is built under:
    ///   languages/openxml/tests/bin/Debug/net10.0/
    /// The main project exe is under:
    ///   languages/openxml/bin/Debug/net10.0/ExcelOps.exe
    /// We walk up from AppContext.BaseDirectory.
    /// </summary>
    private static string FindExe()
    {
        // The test base dir is something like:
        //   …/languages/openxml/tests/bin/Debug/net10.0/
        // Walk up to find ExcelOps.exe relative to the openxml folder.
        string baseDir = AppContext.BaseDirectory;
        // Go up 4 levels from tests/bin/Debug/net10.0 → languages/openxml
        DirectoryInfo? dir = new DirectoryInfo(baseDir);
        for (int i = 0; i < 4; i++) dir = dir?.Parent;

        if (dir is null)
            throw new InvalidOperationException($"Could not navigate up from {baseDir}");

        string candidate = Path.Combine(dir.FullName, "bin", "Debug", "net10.0", "ExcelOps.exe");
        if (!File.Exists(candidate))
            throw new FileNotFoundException($"ExcelOps.exe not found at {candidate}");

        return candidate;
    }

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
        Path.Combine(Path.GetTempPath(), $"ExcelOpsTest_{name}_{Guid.NewGuid():N}.xlsx");

    // ---------------------------------------------------------------------------
    // Tests
    // ---------------------------------------------------------------------------

    [Fact]
    public void WriteString_ok_and_readback()
    {
        string path = TempXlsx("string");
        try
        {
            string req = JsonSerializer.Serialize(new
            {
                op = "cell.write",
                path,
                target = new { sheet = "Sheet1", range = "A1" },
                @params = new { kind = "string", value = "hi" },
            });

            string raw = RunOp(req);

            Assert.Contains("\"ok\":true", raw);

            string? val = OpenXmlReadback.ReadCellString(path, "Sheet1", "A1");
            Assert.Equal("hi", val);
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public void WriteNumber_ok_and_readback()
    {
        string path = TempXlsx("number");
        try
        {
            string req = JsonSerializer.Serialize(new
            {
                op = "cell.write",
                path,
                target = new { sheet = "Sheet1", range = "A2" },
                @params = new { kind = "number", value = 42.5 },
            });

            string raw = RunOp(req);

            Assert.Contains("\"ok\":true", raw);

            string? val = OpenXmlReadback.ReadCellString(path, "Sheet1", "A2");
            // Stored as invariant decimal; "R" round-trip format for 42.5 → "42.5"
            Assert.Equal("42.5", val);
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public void WriteBool_ok_and_readback()
    {
        string path = TempXlsx("bool");
        try
        {
            string req = JsonSerializer.Serialize(new
            {
                op = "cell.write",
                path,
                target = new { sheet = "Sheet1", range = "A3" },
                @params = new { kind = "bool", value = true },
            });

            string raw = RunOp(req);

            Assert.Contains("\"ok\":true", raw);

            string? val = OpenXmlReadback.ReadCellString(path, "Sheet1", "A3");
            // OpenXML stores bool as "1" (true) or "0" (false)
            Assert.Equal("1", val);
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    // ---------------------------------------------------------------------------
    // Extra guard: A1 ordering — writing cells out of order must not corrupt file
    // ---------------------------------------------------------------------------

    [Fact]
    public void EnsureCell_outOfOrder_doesNotCorrupt()
    {
        string path = TempXlsx("order");
        try
        {
            // Write C3 first, then A1 — would corrupt if ordering not enforced
            string req1 = JsonSerializer.Serialize(new
            {
                op = "cell.write",
                path,
                target = new { sheet = "Sheet1", range = "C3" },
                @params = new { kind = "string", value = "first" },
            });
            string req2 = JsonSerializer.Serialize(new
            {
                op = "cell.write",
                path,
                target = new { sheet = "Sheet1", range = "A1" },
                @params = new { kind = "string", value = "second" },
            });

            string r1 = RunOp(req1);
            string r2 = RunOp(req2);

            Assert.Contains("\"ok\":true", r1);
            Assert.Contains("\"ok\":true", r2);

            // Both cells should be readable
            Assert.Equal("first",  OpenXmlReadback.ReadCellString(path, "Sheet1", "C3"));
            Assert.Equal("second", OpenXmlReadback.ReadCellString(path, "Sheet1", "A1"));
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }
}
