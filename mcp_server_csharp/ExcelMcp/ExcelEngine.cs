using System;
using System.CodeDom.Compiler;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using Microsoft.CSharp;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelMcp
{
    /// <summary>
    /// Excel COM engine + CSharpCodeProvider-based CodeAct executor.
    /// All operations run in the same process as Excel (when used as add-in)
    /// or cross-process via COM (when used standalone).
    /// </summary>
    class ExcelEngine
    {
        private Excel.Application _app;
        private Excel.Workbook _wb;
        public string FilePath { get; private set; }

        public ExcelEngine()
        {
            _app = new Excel.Application();
            _app.Visible = true;
            _app.DisplayAlerts = false;
        }

        public void Open(string path)
        {
            FilePath = Path.GetFullPath(path);
            _wb = _app.Workbooks.Open(FilePath);
        }

        public void Create(string path)
        {
            FilePath = Path.GetFullPath(path);
            _wb = _app.Workbooks.Add();
            _wb.SaveAs(FilePath, 51); // xlOpenXMLWorkbook
        }

        public void Save(string path = null)
        {
            if (!string.IsNullOrEmpty(path))
                _wb.SaveAs(Path.GetFullPath(path), 51);
            else
                _wb.Save();
        }

        public void Close()
        {
            try { _wb?.Close(false); } catch { }
            try { _app?.Quit(); } catch { }
            if (_app != null) Marshal.ReleaseComObject(_app);
            _app = null;
        }

        private Excel.Workbook Wb()
        {
            if (_wb != null) return _wb;
            try { _wb = _app.ActiveWorkbook; } catch { }
            if (_wb == null) throw new InvalidOperationException("No workbook open. Use excel_open first.");
            return _wb;
        }

        private Excel.Worksheet GetSheet(string sheetRef)
        {
            Excel.Workbook wb = Wb();
            if (int.TryParse(sheetRef, out int idx))
                return (Excel.Worksheet)wb.Worksheets[idx];
            return (Excel.Worksheet)wb.Worksheets[sheetRef];
        }

        // ---- Inspect ----

        public object Inspect()
        {
            Excel.Workbook wb = Wb();
            var sheets = new List<object>();
            for (int i = 1; i <= wb.Worksheets.Count; i++)
            {
                Excel.Worksheet ws = (Excel.Worksheet)wb.Worksheets[i];
                Excel.Range used = ws.UsedRange;
                sheets.Add(new Dictionary<string, object>
                {
                    ["index"] = i,
                    ["name"] = ws.Name,
                    ["used_range"] = used?.Address ?? "",
                    ["rows"] = used?.Rows.Count ?? 0,
                    ["cols"] = used?.Columns.Count ?? 0,
                });
            }
            return new Dictionary<string, object> { ["sheets"] = sheets };
        }

        public object InspectSheet(string sheetRef, int maxRows, int maxCols)
        {
            Excel.Worksheet ws = GetSheet(sheetRef);
            Excel.Range used = ws.UsedRange;
            int rows = Math.Min(used?.Rows.Count ?? 0, maxRows);
            int cols = Math.Min(used?.Columns.Count ?? 0, maxCols);
            int startRow = used?.Row ?? 1;
            int startCol = used?.Column ?? 1;

            var data = new List<List<object>>();
            for (int r = startRow; r < startRow + rows; r++)
            {
                var row = new List<object>();
                for (int c = startCol; c < startCol + cols; c++)
                    row.Add(((Excel.Range)ws.Cells[r, c]).Value2);
                data.Add(row);
            }
            return new Dictionary<string, object>
            {
                ["name"] = ws.Name,
                ["used_range"] = used?.Address ?? "",
                ["data"] = data
            };
        }

        // ---- CodeAct: CSharpCodeProvider ----

        public string ExecuteCode(string code)
        {
            Excel.Workbook wb = Wb();
            Excel.Worksheet ws = (Excel.Worksheet)wb.ActiveSheet;

            // Wrap user code in a static method
            string fullSource = @"
using System;
using System.Collections.Generic;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

public class ExcelScript
{
    public static string Execute(Excel.Application app, Excel.Workbook wb, Excel.Worksheet ws)
    {
" + code + @"
        return null;
    }
}";

            // Compile
            var provider = new CSharpCodeProvider();
            var parameters = new CompilerParameters
            {
                GenerateInMemory = true,
                GenerateExecutable = false,
                TreatWarningsAsErrors = false,
            };

            // Add references
            parameters.ReferencedAssemblies.Add("System.dll");
            parameters.ReferencedAssemblies.Add("System.Core.dll");
            parameters.ReferencedAssemblies.Add("Microsoft.CSharp.dll");
            parameters.ReferencedAssemblies.Add(typeof(Excel.Application).Assembly.Location);
            parameters.ReferencedAssemblies.Add(typeof(Marshal).Assembly.Location);

            CompilerResults results = provider.CompileAssemblyFromSource(parameters, fullSource);

            if (results.Errors.HasErrors)
            {
                var sb = new StringBuilder();
                foreach (CompilerError err in results.Errors)
                {
                    if (!err.IsWarning)
                        sb.AppendLine($"Line {err.Line - 9}: {err.ErrorText}");
                }
                return JsonConvert.SerializeObject(new
                {
                    ok = false,
                    error = "Compilation failed",
                    details = sb.ToString()
                });
            }

            // Execute
            try
            {
                Assembly assembly = results.CompiledAssembly;
                Type scriptType = assembly.GetType("ExcelScript");
                MethodInfo method = scriptType.GetMethod("Execute");
                object result = method.Invoke(null, new object[] { _app, wb, ws });
                return JsonConvert.SerializeObject(new
                {
                    ok = true,
                    result = result?.ToString()
                });
            }
            catch (TargetInvocationException ex)
            {
                Exception inner = ex.InnerException ?? ex;
                return JsonConvert.SerializeObject(new
                {
                    ok = false,
                    error = inner.Message,
                    traceback = inner.StackTrace
                });
            }
            catch (Exception ex)
            {
                return JsonConvert.SerializeObject(new
                {
                    ok = false,
                    error = ex.Message,
                    traceback = ex.StackTrace
                });
            }
        }

        // ---- Action dispatch (structured JSON) ----

        public string ExecuteActions(string actionsJson)
        {
            JToken parsed = JToken.Parse(actionsJson);
            JArray actions = parsed is JArray arr ? arr : new JArray { parsed };
            var results = new List<object>();

            foreach (JObject action in actions)
            {
                try
                {
                    results.Add(DispatchAction(action));
                }
                catch (Exception ex)
                {
                    results.Add(new Dictionary<string, object>
                    {
                        ["ok"] = false,
                        ["action"] = (string)action["action"] ?? "",
                        ["error"] = ex.Message
                    });
                }
            }

            return JsonConvert.SerializeObject(results);
        }

        private object DispatchAction(JObject action)
        {
            string name = (string)action["action"] ?? "";
            string sheet = action["sheet"]?.ToString() ?? "1";
            JObject pars = action["params"] as JObject ?? new JObject();

            switch (name)
            {
                case "write_cell":
                {
                    Excel.Worksheet ws = GetSheet(sheet);
                    string cell = (string)action["cell"];
                    string kind = (string)action["kind"] ?? "auto";
                    if (kind == "formula")
                        ws.Range[cell].Formula = action["value"]?.ToString();
                    else
                        ws.Range[cell].Value2 = action["value"]?.ToObject<object>();
                    return Ok(name);
                }
                case "read_cell":
                {
                    Excel.Worksheet ws = GetSheet(sheet);
                    string cell = (string)action["cell"];
                    string prop = (string)action["property"] ?? "value2";
                    object val;
                    if (prop == "formula") val = ws.Range[cell].Formula;
                    else if (prop == "text") val = ws.Range[cell].Text;
                    else val = ws.Range[cell].Value2;
                    return new Dictionary<string, object> { ["ok"] = true, ["action"] = name, ["value"] = val };
                }
                case "write_range":
                {
                    Excel.Worksheet ws = GetSheet(sheet);
                    JArray rows = action["values"] as JArray;
                    if (rows != null)
                    {
                        int rc = rows.Count;
                        int cc = (rows[0] as JArray)?.Count ?? 1;
                        object[,] data = new object[rc, cc];
                        for (int r = 0; r < rc; r++)
                        {
                            JArray row = rows[r] as JArray;
                            for (int c = 0; c < cc; c++)
                                data[r, c] = row?[c]?.ToObject<object>();
                        }
                        ws.Range[(string)action["range"]].Value2 = data;
                    }
                    return Ok(name);
                }
                case "read_range":
                {
                    object val = GetSheet(sheet).Range[(string)action["range"]].Value2;
                    return new Dictionary<string, object> { ["ok"] = true, ["action"] = name, ["values"] = val };
                }
                case "clear_range":
                {
                    string mode = (string)pars["mode"] ?? "contents";
                    Excel.Range rng = GetSheet(sheet).Range[(string)action["range"]];
                    if (mode == "all") rng.Clear(); else rng.ClearContents();
                    return Ok(name);
                }
                case "merge_cells":
                    GetSheet(sheet).Range[(string)action["range"]].Merge();
                    return Ok(name);
                case "unmerge_cells":
                    GetSheet(sheet).Range[(string)action["range"]].UnMerge();
                    return Ok(name);
                case "set_format":
                {
                    Excel.Range rng = GetSheet(sheet).Range[(string)action["range"]];
                    if (pars["bold"] != null) rng.Font.Bold = (bool)pars["bold"];
                    if (pars["italic"] != null) rng.Font.Italic = (bool)pars["italic"];
                    if (pars["font_size"] != null) rng.Font.Size = (double)pars["font_size"];
                    if (pars["font_name"] != null) rng.Font.Name = (string)pars["font_name"];
                    if (pars["font_color"] != null) rng.Font.Color = (int)pars["font_color"];
                    if (pars["bg_color"] != null) rng.Interior.Color = (int)pars["bg_color"];
                    if (pars["number_format"] != null) rng.NumberFormat = (string)pars["number_format"];
                    return Ok(name);
                }
                case "insert_rows":
                {
                    Excel.Worksheet ws = GetSheet(sheet);
                    int row = (int)action["row"];
                    int count = pars["count"]?.Value<int>() ?? 1;
                    for (int i = 0; i < count; i++) ((Excel.Range)ws.Rows[row]).Insert();
                    return Ok(name);
                }
                case "delete_rows":
                {
                    Excel.Worksheet ws = GetSheet(sheet);
                    int row = (int)action["row"];
                    int count = pars["count"]?.Value<int>() ?? 1;
                    ((Excel.Range)ws.Rows[row + ":" + (row + count - 1)]).Delete();
                    return Ok(name);
                }
                case "add_sheet":
                {
                    Excel.Worksheet ws = (Excel.Worksheet)Wb().Worksheets.Add();
                    string sname = (string)pars["name"];
                    if (!string.IsNullOrEmpty(sname)) ws.Name = sname;
                    return new Dictionary<string, object> { ["ok"] = true, ["action"] = name, ["name"] = ws.Name };
                }
                case "rename_sheet":
                    ((Excel.Worksheet)Wb().Worksheets[(string)action["old_name"]]).Name = (string)action["new_name"];
                    return Ok(name);
                case "delete_sheet":
                    GetSheet(sheet).Delete();
                    return Ok(name);
                case "sort_range":
                {
                    Excel.Worksheet ws = GetSheet(sheet);
                    Excel.Range rng = ws.Range[(string)action["range"]];
                    string order = (string)pars["order"] ?? "asc";
                    rng.Sort(Key1: ws.Range[(string)action["key_col"]],
                             Order1: order == "desc" ? Excel.XlSortOrder.xlDescending : Excel.XlSortOrder.xlAscending,
                             Header: Excel.XlYesNoGuess.xlYes);
                    return Ok(name);
                }
                case "find_replace":
                    GetSheet(sheet).Range[(string)action["range"]].Replace(
                        What: (string)action["find"], Replacement: (string)action["replace"]);
                    return Ok(name);
                case "calculate":
                    _app.Calculate();
                    return Ok(name);
                case "export_pdf":
                    Wb().ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, (string)action["path"]);
                    return Ok(name);
                default:
                    return new Dictionary<string, object> { ["ok"] = false, ["error"] = "Unknown action: " + name };
            }
        }

        private static Dictionary<string, object> Ok(string action)
        {
            return new Dictionary<string, object> { ["ok"] = true, ["action"] = action };
        }
    }
}
