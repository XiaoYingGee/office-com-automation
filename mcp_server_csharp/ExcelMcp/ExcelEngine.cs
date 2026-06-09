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
    class ExcelEngine
    {
        private Excel.Application _app;
        private Excel.Workbook _wb;
        private string _filePath;

        public string FilePath { get { return _filePath; } }

        public ExcelEngine()
        {
            _app = new Excel.Application();
            _app.Visible = true;
            _app.DisplayAlerts = false;
        }

        public void Open(string path)
        {
            _filePath = Path.GetFullPath(path);
            _wb = _app.Workbooks.Open(_filePath);
        }

        public void Create(string path)
        {
            _filePath = Path.GetFullPath(path);
            _wb = _app.Workbooks.Add();
            _wb.SaveAs(_filePath, 51);
        }

        public void Save(string path)
        {
            if (!string.IsNullOrEmpty(path))
                _wb.SaveAs(Path.GetFullPath(path), 51);
            else
                _wb.Save();
        }

        public void Close()
        {
            try { if (_wb != null) _wb.Close(false); } catch { }
            try { if (_app != null) _app.Quit(); } catch { }
            if (_app != null) Marshal.ReleaseComObject(_app);
            _app = null;
        }

        private Excel.Workbook Wb()
        {
            if (_wb != null) return _wb;
            try { _wb = _app.ActiveWorkbook; } catch { }
            if (_wb == null) throw new InvalidOperationException("No workbook open.");
            return _wb;
        }

        private Excel.Worksheet GetSheet(string sheetRef)
        {
            Excel.Workbook wb = Wb();
            int idx;
            if (int.TryParse(sheetRef, out idx))
                return (Excel.Worksheet)wb.Worksheets[idx];
            return (Excel.Worksheet)wb.Worksheets[sheetRef];
        }

        // ---- Inspect ----

        public string Inspect()
        {
            Excel.Workbook wb = Wb();
            var sheets = new List<Dictionary<string, object>>();
            for (int i = 1; i <= wb.Worksheets.Count; i++)
            {
                Excel.Worksheet ws = (Excel.Worksheet)wb.Worksheets[i];
                Excel.Range used = ws.UsedRange;
                var d = new Dictionary<string, object>();
                d.Add("index", i);
                d.Add("name", ws.Name);
                d.Add("used_range", used != null ? used.Address : "");
                d.Add("rows", used != null ? used.Rows.Count : 0);
                d.Add("cols", used != null ? used.Columns.Count : 0);
                sheets.Add(d);
            }
            var result = new Dictionary<string, object>();
            result.Add("sheets", sheets);
            return JsonConvert.SerializeObject(result);
        }

        public string InspectSheet(string sheetRef, int maxRows, int maxCols)
        {
            Excel.Worksheet ws = GetSheet(sheetRef);
            Excel.Range used = ws.UsedRange;
            int rows = Math.Min(used != null ? used.Rows.Count : 0, maxRows);
            int cols = Math.Min(used != null ? used.Columns.Count : 0, maxCols);
            int startRow = used != null ? used.Row : 1;
            int startCol = used != null ? used.Column : 1;

            var data = new List<List<object>>();
            for (int r = startRow; r < startRow + rows; r++)
            {
                var row = new List<object>();
                for (int c = startCol; c < startCol + cols; c++)
                    row.Add(((Excel.Range)ws.Cells[r, c]).Value2);
                data.Add(row);
            }
            var result = new Dictionary<string, object>();
            result.Add("name", ws.Name);
            result.Add("used_range", used != null ? used.Address : "");
            result.Add("data", data);
            return JsonConvert.SerializeObject(result);
        }

        // ---- CodeAct: CSharpCodeProvider ----

        public string ExecuteCode(string code)
        {
            Excel.Workbook wb = Wb();
            Excel.Worksheet ws = (Excel.Worksheet)wb.ActiveSheet;

            string fullSource = @"
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

public class ExcelScript
{
    public static string Execute(Excel.Application app, Excel.Workbook wb, Excel.Worksheet ws)
    {
" + code + @"
        return null;
    }
}";
            var provider = new CSharpCodeProvider();
            var parameters = new CompilerParameters();
            parameters.GenerateInMemory = true;
            parameters.GenerateExecutable = false;
            parameters.TreatWarningsAsErrors = false;
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
                        sb.AppendLine("Line " + (err.Line - 9) + ": " + err.ErrorText);
                }
                var errResult = new Dictionary<string, object>();
                errResult.Add("ok", false);
                errResult.Add("error", "Compilation failed");
                errResult.Add("details", sb.ToString());
                return JsonConvert.SerializeObject(errResult);
            }

            try
            {
                Assembly assembly = results.CompiledAssembly;
                Type scriptType = assembly.GetType("ExcelScript");
                MethodInfo method = scriptType.GetMethod("Execute");
                object result = method.Invoke(null, new object[] { _app, wb, ws });
                var okResult = new Dictionary<string, object>();
                okResult.Add("ok", true);
                okResult.Add("result", result != null ? result.ToString() : null);
                return JsonConvert.SerializeObject(okResult);
            }
            catch (TargetInvocationException ex)
            {
                Exception inner = ex.InnerException != null ? ex.InnerException : ex;
                var errResult = new Dictionary<string, object>();
                errResult.Add("ok", false);
                errResult.Add("error", inner.Message);
                errResult.Add("traceback", inner.StackTrace);
                return JsonConvert.SerializeObject(errResult);
            }
            catch (Exception ex)
            {
                var errResult = new Dictionary<string, object>();
                errResult.Add("ok", false);
                errResult.Add("error", ex.Message);
                errResult.Add("traceback", ex.StackTrace);
                return JsonConvert.SerializeObject(errResult);
            }
        }

        // ---- Action dispatch ----

        public string ExecuteActions(string actionsJson)
        {
            JToken parsed = JToken.Parse(actionsJson);
            JArray actions = parsed is JArray ? (JArray)parsed : new JArray { parsed };
            var results = new List<object>();

            foreach (JToken item in actions)
            {
                JObject action = (JObject)item;
                try
                {
                    results.Add(DispatchAction(action));
                }
                catch (Exception ex)
                {
                    var err = new Dictionary<string, object>();
                    err.Add("ok", false);
                    err.Add("action", (string)action["action"] ?? "");
                    err.Add("error", ex.Message);
                    results.Add(err);
                }
            }

            return JsonConvert.SerializeObject(results);
        }

        private Dictionary<string, object> DispatchAction(JObject action)
        {
            string name = (string)action["action"] ?? "";
            string sheet = action["sheet"] != null ? action["sheet"].ToString() : "1";
            JObject pars = action["params"] as JObject;
            if (pars == null) pars = new JObject();

            switch (name)
            {
                case "write_cell":
                {
                    Excel.Worksheet ws = GetSheet(sheet);
                    string cell = (string)action["cell"];
                    string kind = (string)action["kind"] ?? "auto";
                    if (kind == "formula")
                        ws.Range[cell].Formula = action["value"] != null ? action["value"].ToString() : "";
                    else
                        ws.Range[cell].Value2 = action["value"] != null ? action["value"].ToObject<object>() : null;
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
                    var r = new Dictionary<string, object>();
                    r.Add("ok", true);
                    r.Add("action", name);
                    r.Add("value", val);
                    return r;
                }
                case "write_range":
                {
                    Excel.Worksheet ws = GetSheet(sheet);
                    JArray rows = action["values"] as JArray;
                    if (rows != null)
                    {
                        int rc = rows.Count;
                        JArray firstRow = rows[0] as JArray;
                        int cc = firstRow != null ? firstRow.Count : 1;
                        object[,] data = new object[rc, cc];
                        for (int ri = 0; ri < rc; ri++)
                        {
                            JArray row = rows[ri] as JArray;
                            for (int ci = 0; ci < cc; ci++)
                                data[ri, ci] = row != null && row[ci] != null ? row[ci].ToObject<object>() : null;
                        }
                        ws.Range[(string)action["range"]].Value2 = data;
                    }
                    return Ok(name);
                }
                case "read_range":
                {
                    object val = GetSheet(sheet).Range[(string)action["range"]].Value2;
                    var r = new Dictionary<string, object>();
                    r.Add("ok", true);
                    r.Add("action", name);
                    r.Add("values", val);
                    return r;
                }
                case "clear_range":
                {
                    string mode = pars["mode"] != null ? (string)pars["mode"] : "contents";
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
                    int count = pars["count"] != null ? pars["count"].Value<int>() : 1;
                    for (int i = 0; i < count; i++) ((Excel.Range)ws.Rows[row]).Insert();
                    return Ok(name);
                }
                case "delete_rows":
                {
                    Excel.Worksheet ws = GetSheet(sheet);
                    int row = (int)action["row"];
                    int count = pars["count"] != null ? pars["count"].Value<int>() : 1;
                    ((Excel.Range)ws.Rows[row + ":" + (row + count - 1)]).Delete();
                    return Ok(name);
                }
                case "add_sheet":
                {
                    Excel.Worksheet ws = (Excel.Worksheet)Wb().Worksheets.Add();
                    string sname = pars["name"] != null ? (string)pars["name"] : null;
                    if (!string.IsNullOrEmpty(sname)) ws.Name = sname;
                    var r = new Dictionary<string, object>();
                    r.Add("ok", true);
                    r.Add("action", name);
                    r.Add("name", ws.Name);
                    return r;
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
                    string order = pars["order"] != null ? (string)pars["order"] : "asc";
                    Excel.XlSortOrder xlOrder = order == "desc" ? Excel.XlSortOrder.xlDescending : Excel.XlSortOrder.xlAscending;
                    rng.Sort(Key1: ws.Range[(string)action["key_col"]], Order1: xlOrder, Header: Excel.XlYesNoGuess.xlYes);
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
                {
                    var r = new Dictionary<string, object>();
                    r.Add("ok", false);
                    r.Add("error", "Unknown action: " + name);
                    return r;
                }
            }
        }

        private static Dictionary<string, object> Ok(string action)
        {
            var r = new Dictionary<string, object>();
            r.Add("ok", true);
            r.Add("action", action);
            return r;
        }
    }
}
