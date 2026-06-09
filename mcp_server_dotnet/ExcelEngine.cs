using System.Diagnostics;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelMcpDotnet;

public class ExcelEngine : IDisposable
{
    private Excel.Application? _app;
    private Excel.Workbook? _wb;
    private string? _filePath;

    public string? FilePath => _filePath;

    private Excel.Application App
    {
        get
        {
            if (_app == null)
            {
                _app = new Excel.Application();
                _app.Visible = true;
                _app.DisplayAlerts = false;
            }
            return _app;
        }
    }

    private Excel.Workbook Wb
    {
        get
        {
            if (_wb != null) return _wb;
            try { _wb = App.ActiveWorkbook; } catch { }
            if (_wb == null) throw new InvalidOperationException("No workbook open. Use excel_open first.");
            return _wb;
        }
    }

    public string Open(string path)
    {
        var sw = Stopwatch.StartNew();
        _filePath = Path.GetFullPath(path);
        _wb = App.Workbooks.Open(_filePath);
        sw.Stop();
        var info = InspectInternal();
        return JsonSerializer.Serialize(new
        {
            ok = true,
            path = _filePath,
            sheets = info.Select(s => s.name).ToArray(),
            elapsed_ms = sw.ElapsedMilliseconds
        });
    }

    public string Create(string path)
    {
        var sw = Stopwatch.StartNew();
        _filePath = Path.GetFullPath(path);
        _wb = App.Workbooks.Add();
        _wb.SaveAs(_filePath, 51);
        sw.Stop();
        return JsonSerializer.Serialize(new
        {
            ok = true,
            path = _filePath,
            elapsed_ms = sw.ElapsedMilliseconds
        });
    }

    public string Save(string? path = null)
    {
        var sw = Stopwatch.StartNew();
        if (!string.IsNullOrEmpty(path))
            Wb.SaveAs(Path.GetFullPath(path), 51);
        else
            Wb.Save();
        sw.Stop();
        return JsonSerializer.Serialize(new { ok = true, elapsed_ms = sw.ElapsedMilliseconds });
    }

    public string Inspect()
    {
        var sheets = InspectInternal();
        return JsonSerializer.Serialize(new { sheets });
    }

    public string InspectSheet(string sheetRef, int maxRows = 50, int maxCols = 26)
    {
        var ws = GetSheet(sheetRef);
        Excel.Range used = ws.UsedRange;
        int rows = Math.Min(used?.Rows.Count ?? 0, maxRows);
        int cols = Math.Min(used?.Columns.Count ?? 0, maxCols);
        int startRow = used?.Row ?? 1;
        int startCol = used?.Column ?? 1;

        var data = new List<List<object?>>();
        for (int r = startRow; r < startRow + rows; r++)
        {
            var row = new List<object?>();
            for (int c = startCol; c < startCol + cols; c++)
                row.Add(((Excel.Range)ws.Cells[r, c]).Value2);
            data.Add(row);
        }
        return JsonSerializer.Serialize(new { name = ws.Name, used_range = used?.Address ?? "", data });
    }

    public string ExecuteCode(string code)
    {
        var wb = Wb;
        var ws = (Excel.Worksheet)wb.ActiveSheet;

        string fullSource = $$"""
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

public class ExcelScript
{
    public static string Execute(Excel.Application app, Excel.Workbook wb, Excel.Worksheet ws)
    {
{{code}}
        return null;
    }
}
""";

        var sw = Stopwatch.StartNew();

        var syntaxTree = CSharpSyntaxTree.ParseText(fullSource);

        var references = new List<MetadataReference>
        {
            MetadataReference.CreateFromFile(typeof(object).Assembly.Location),
            MetadataReference.CreateFromFile(typeof(Console).Assembly.Location),
            MetadataReference.CreateFromFile(typeof(Enumerable).Assembly.Location),
            MetadataReference.CreateFromFile(typeof(Marshal).Assembly.Location),
            MetadataReference.CreateFromFile(typeof(Excel.Application).Assembly.Location),
            MetadataReference.CreateFromFile(typeof(Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo).Assembly.Location),
        };

        // Add runtime assemblies
        var runtimeDir = Path.GetDirectoryName(typeof(object).Assembly.Location)!;
        references.Add(MetadataReference.CreateFromFile(Path.Combine(runtimeDir, "System.Runtime.dll")));
        references.Add(MetadataReference.CreateFromFile(Path.Combine(runtimeDir, "System.Collections.dll")));
        references.Add(MetadataReference.CreateFromFile(Path.Combine(runtimeDir, "System.Linq.dll")));
        references.Add(MetadataReference.CreateFromFile(Path.Combine(runtimeDir, "netstandard.dll")));

        var compilation = CSharpCompilation.Create("ExcelScript_" + Guid.NewGuid().ToString("N"))
            .WithOptions(new CSharpCompilationOptions(OutputKind.DynamicallyLinkedLibrary))
            .AddReferences(references)
            .AddSyntaxTrees(syntaxTree);

        using var ms = new MemoryStream();
        var emitResult = compilation.Emit(ms);

        if (!emitResult.Success)
        {
            sw.Stop();
            var errors = emitResult.Diagnostics
                .Where(d => d.Severity == DiagnosticSeverity.Error)
                .Select(d =>
                {
                    var lineSpan = d.Location.GetLineSpan();
                    int line = lineSpan.StartLinePosition.Line - 9; // offset for wrapper
                    return $"Line {line}: {d.GetMessage()}";
                });
            return JsonSerializer.Serialize(new
            {
                ok = false,
                error = "Compilation failed",
                details = string.Join("\n", errors),
                elapsed_ms = sw.ElapsedMilliseconds
            });
        }

        try
        {
            ms.Seek(0, SeekOrigin.Begin);
            var assembly = System.Runtime.Loader.AssemblyLoadContext.Default.LoadFromStream(ms);
            var scriptType = assembly.GetType("ExcelScript")!;
            var method = scriptType.GetMethod("Execute")!;
            var result = method.Invoke(null, [_app, wb, ws]);
            sw.Stop();
            return JsonSerializer.Serialize(new
            {
                ok = true,
                result = result?.ToString(),
                elapsed_ms = sw.ElapsedMilliseconds
            });
        }
        catch (TargetInvocationException ex)
        {
            sw.Stop();
            var inner = ex.InnerException ?? ex;
            return JsonSerializer.Serialize(new
            {
                ok = false,
                error = inner.Message,
                traceback = inner.StackTrace,
                elapsed_ms = sw.ElapsedMilliseconds
            });
        }
        catch (Exception ex)
        {
            sw.Stop();
            return JsonSerializer.Serialize(new
            {
                ok = false,
                error = ex.Message,
                traceback = ex.StackTrace,
                elapsed_ms = sw.ElapsedMilliseconds
            });
        }
    }

    public string ExecuteActions(string actionsJson)
    {
        using var doc = JsonDocument.Parse(actionsJson);
        var actions = doc.RootElement.ValueKind == JsonValueKind.Array
            ? doc.RootElement.EnumerateArray().ToList()
            : [doc.RootElement];

        var results = new List<object>();
        foreach (var action in actions)
        {
            try
            {
                results.Add(DispatchAction(action));
            }
            catch (Exception ex)
            {
                results.Add(new { ok = false, action = action.GetProperty("action").GetString() ?? "", error = ex.Message });
            }
        }
        return JsonSerializer.Serialize(results);
    }

    private object DispatchAction(JsonElement action)
    {
        string name = action.GetProperty("action").GetString() ?? "";
        string sheet = action.TryGetProperty("sheet", out var s) ? s.GetString() ?? "1" : "1";
        var ws = GetSheet(sheet);

        switch (name)
        {
            case "write_cell":
            {
                string cell = action.GetProperty("cell").GetString()!;
                var val = action.GetProperty("value");
                string kind = action.TryGetProperty("kind", out var k) ? k.GetString() ?? "auto" : "auto";
                if (kind == "formula")
                    ws.Range[cell].Formula = val.GetString();
                else
                    ws.Range[cell].Value2 = JsonElementToObject(val);
                return new { ok = true, action = name };
            }
            case "read_cell":
            {
                string cell = action.GetProperty("cell").GetString()!;
                object? val = ws.Range[cell].Value2;
                return new { ok = true, action = name, value = val };
            }
            case "write_range":
            {
                string range = action.GetProperty("range").GetString()!;
                if (action.TryGetProperty("values", out var rows))
                {
                    int rc = rows.GetArrayLength();
                    int cc = rows[0].GetArrayLength();
                    var data = new object?[rc, cc];
                    for (int ri = 0; ri < rc; ri++)
                        for (int ci = 0; ci < cc; ci++)
                            data[ri, ci] = JsonElementToObject(rows[ri][ci]);
                    ws.Range[range].Value2 = data;
                }
                return new { ok = true, action = name };
            }
            case "read_range":
            {
                string range = action.GetProperty("range").GetString()!;
                object? val = ws.Range[range].Value2;
                return new { ok = true, action = name, values = val };
            }
            case "clear_range":
            {
                string range = action.GetProperty("range").GetString()!;
                ws.Range[range].ClearContents();
                return new { ok = true, action = name };
            }
            case "merge_cells":
                ws.Range[action.GetProperty("range").GetString()!].Merge();
                return new { ok = true, action = name };
            case "set_format":
            {
                var rng = ws.Range[action.GetProperty("range").GetString()!];
                if (action.TryGetProperty("params", out var p))
                {
                    if (p.TryGetProperty("bold", out var b)) rng.Font.Bold = b.GetBoolean();
                    if (p.TryGetProperty("font_size", out var fs)) rng.Font.Size = fs.GetDouble();
                    if (p.TryGetProperty("font_color", out var fc)) rng.Font.Color = fc.GetInt32();
                    if (p.TryGetProperty("bg_color", out var bg)) rng.Interior.Color = bg.GetInt32();
                    if (p.TryGetProperty("number_format", out var nf)) rng.NumberFormat = nf.GetString();
                }
                return new { ok = true, action = name };
            }
            case "add_sheet":
            {
                var newWs = (Excel.Worksheet)Wb.Worksheets.Add();
                if (action.TryGetProperty("params", out var p) && p.TryGetProperty("name", out var n))
                    newWs.Name = n.GetString()!;
                return new { ok = true, action = name, name2 = newWs.Name };
            }
            case "delete_sheet":
                ws.Delete();
                return new { ok = true, action = name };
            default:
                return new { ok = false, error = $"Unknown action: {name}" };
        }
    }

    private List<(string name, string usedRange, int rows, int cols)> InspectInternal()
    {
        var wb = Wb;
        var sheets = new List<(string name, string usedRange, int rows, int cols)>();
        for (int i = 1; i <= wb.Worksheets.Count; i++)
        {
            var ws = (Excel.Worksheet)wb.Worksheets[i];
            var used = ws.UsedRange;
            sheets.Add((ws.Name, used?.Address ?? "", used?.Rows.Count ?? 0, used?.Columns.Count ?? 0));
        }
        return sheets;
    }

    private Excel.Worksheet GetSheet(string sheetRef)
    {
        if (int.TryParse(sheetRef, out int idx))
            return (Excel.Worksheet)Wb.Worksheets[idx];
        return (Excel.Worksheet)Wb.Worksheets[sheetRef];
    }

    private static object? JsonElementToObject(JsonElement el) => el.ValueKind switch
    {
        JsonValueKind.Number => el.TryGetInt64(out long l) ? l : el.GetDouble(),
        JsonValueKind.String => el.GetString(),
        JsonValueKind.True => true,
        JsonValueKind.False => false,
        JsonValueKind.Null => null,
        _ => el.GetRawText()
    };

    public void Dispose()
    {
        try { _wb?.Close(false); } catch { }
        try { _app?.Quit(); } catch { }
        if (_app != null) Marshal.ReleaseComObject(_app);
        _app = null;
        _wb = null;
    }
}
