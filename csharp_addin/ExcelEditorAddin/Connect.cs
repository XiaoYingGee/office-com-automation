using System;
using System.Collections.Generic;
using System.Globalization;
using System.Runtime.InteropServices;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Office = Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelEditorAddin
{
    [Guid("289E9AF1-4973-11D1-AE81-00A0C90F26F4")]
    public enum ext_ConnectMode
    {
        ext_cm_AfterStartup = 0,
        ext_cm_Startup = 1,
        ext_cm_External = 2,
        ext_cm_CommandLine = 3,
        ext_cm_Solution = 4,
        ext_cm_UISetup = 5,
    }

    [Guid("289E9AF2-4973-11D1-AE81-00A0C90F26F4")]
    public enum ext_DisconnectMode
    {
        ext_dm_HostShutdown = 0,
        ext_dm_UserClosed = 1,
        ext_dm_UISetupComplete = 2,
        ext_dm_SolutionClosed = 3,
    }

    [ComImport]
    [Guid("B65AD801-ABAF-11D0-BB8B-00A0C90F2744")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IDTExtensibility2
    {
        [DispId(1)]
        void OnConnection(
            [In, MarshalAs(UnmanagedType.IDispatch)] object Application,
            [In] ext_ConnectMode ConnectMode,
            [In, MarshalAs(UnmanagedType.IDispatch)] object AddInInst,
            [In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);

        [DispId(2)]
        void OnDisconnection(
            [In] ext_DisconnectMode RemoveMode,
            [In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);

        [DispId(3)]
        void OnAddInsUpdate(
            [In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);

        [DispId(4)]
        void OnStartupComplete(
            [In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);

        [DispId(5)]
        void OnBeginShutdown(
            [In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);
    }

    [ComVisible(true)]
    [Guid("A1B2C3D4-E5F6-4789-ABCD-EF0123456789")]
    [ProgId("ExcelEditor.AddIn")]
#pragma warning disable CS0618
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
#pragma warning restore CS0618
    public class Connect : IDTExtensibility2
    {
        private Excel.Application _app;

        private static void Log(string msg)
        {
            try
            {
                System.IO.File.AppendAllText(
                    System.IO.Path.Combine(System.IO.Path.GetTempPath(), "exceleditor_addin.log"),
                    DateTime.Now.ToString("HH:mm:ss.fff") + " " + msg + Environment.NewLine);
            }
            catch { }
        }

        // ---- IDTExtensibility2 ----

        public void OnConnection(object application, ext_ConnectMode connectMode, object addInInst, ref Array custom)
        {
            Log("OnConnection enter connectMode=" + connectMode);
            _app = (Excel.Application)application;
            try
            {
                ((Office.COMAddIn)addInInst).Object = this;
                Log("OnConnection set addInInst.Object OK");
            }
            catch (Exception ex)
            {
                Log("OnConnection set Object FAILED: " + ex.Message);
            }
        }

        public void OnDisconnection(ext_DisconnectMode removeMode, ref Array custom)
        {
            _app = null;
        }

        public void OnAddInsUpdate(ref Array custom) { }
        public void OnStartupComplete(ref Array custom) { }
        public void OnBeginShutdown(ref Array custom) { }

        public object RequestComAddInAutomationService() => this;

        // ---- Bridge surface ----

        public string Ping() => "pong";

        public string InspectJson()
        {
            return JsonConvert.SerializeObject(InspectWorkbook());
        }

        public string ExecuteActionJson(string actionJson)
        {
            JObject action = JObject.Parse(actionJson);
            return JsonConvert.SerializeObject(ExecuteAction(action));
        }

        // ---- Workbook resolution ----

        private Excel.Workbook Wb()
        {
            try { return _app.ActiveWorkbook; }
            catch { }
            return _app.Workbooks[1];
        }

        // ---- Inspect ----

        private object InspectWorkbook()
        {
            Excel.Workbook wb = Wb();
            var sheets = new List<object>();
            int count = wb.Worksheets.Count;
            for (int i = 1; i <= count; i++)
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

        // ---- Action dispatch ----

        private object ExecuteAction(JObject action)
        {
            string name = (string)action["action"] ?? "";
            string sheet = action["sheet"]?.ToString() ?? "1";
            JObject pars = action["params"] as JObject ?? new JObject();

            switch (name)
            {
                case "write_cell": return WriteCell(action, sheet);
                case "read_cell": return ReadCell(action, sheet);
                case "write_range": return WriteRange(action, sheet);
                case "read_range": return ReadRange(action, sheet);
                case "clear_range": return ClearRange(action, sheet, pars);
                case "merge_cells": return MergeCells(action, sheet);
                case "unmerge_cells": return UnmergeCells(action, sheet);
                case "set_format": return SetFormat(action, sheet, pars);
                case "insert_rows": return InsertRows(action, sheet, pars);
                case "delete_rows": return DeleteRows(action, sheet, pars);
                case "insert_cols": return InsertCols(action, sheet, pars);
                case "delete_cols": return DeleteCols(action, sheet, pars);
                case "add_sheet": return AddSheet(pars);
                case "rename_sheet": return RenameSheet(action);
                case "delete_sheet": return DeleteSheet(sheet);
                case "sort_range": return SortRange(action, sheet, pars);
                case "find_replace": return FindReplace(action, sheet);
                case "calculate": _app.Calculate(); return Ok(name);
                case "export_pdf": return ExportPdf(action);
                default: return new Dictionary<string, object> { ["ok"] = false, ["error"] = "Unknown action: " + name };
            }
        }

        private Excel.Worksheet GetSheet(string sheetRef)
        {
            Excel.Workbook wb = Wb();
            if (int.TryParse(sheetRef, out int idx))
                return (Excel.Worksheet)wb.Worksheets[idx];
            return (Excel.Worksheet)wb.Worksheets[sheetRef];
        }

        // ---- Actions ----

        private object WriteCell(JObject a, string sheet)
        {
            Excel.Worksheet ws = GetSheet(sheet);
            string cell = (string)a["cell"];
            Excel.Range rng = ws.Range[cell];
            string kind = (string)a["kind"] ?? "auto";
            if (kind == "formula")
                rng.Formula = a["value"]?.ToString();
            else
                rng.Value2 = a["value"]?.ToObject<object>();
            return Ok("write_cell");
        }

        private object ReadCell(JObject a, string sheet)
        {
            Excel.Worksheet ws = GetSheet(sheet);
            string cell = (string)a["cell"];
            string prop = (string)a["property"] ?? "value2";
            Excel.Range rng = ws.Range[cell];
            object val;
            if (prop == "formula") val = rng.Formula;
            else if (prop == "text") val = rng.Text;
            else val = rng.Value2;
            return new Dictionary<string, object> { ["ok"] = true, ["action"] = "read_cell", ["value"] = val };
        }

        private object WriteRange(JObject a, string sheet)
        {
            Excel.Worksheet ws = GetSheet(sheet);
            string addr = (string)a["range"];
            JArray rows = a["values"] as JArray;
            if (rows == null) return Ok("write_range");

            int rowCount = rows.Count;
            int colCount = (rows[0] as JArray)?.Count ?? 1;
            object[,] data = new object[rowCount, colCount];
            for (int r = 0; r < rowCount; r++)
            {
                JArray row = rows[r] as JArray;
                for (int c = 0; c < colCount; c++)
                    data[r, c] = row?[c]?.ToObject<object>();
            }
            ws.Range[addr].Value2 = data;
            return Ok("write_range");
        }

        private object ReadRange(JObject a, string sheet)
        {
            Excel.Worksheet ws = GetSheet(sheet);
            string addr = (string)a["range"];
            object val = ws.Range[addr].Value2;
            return new Dictionary<string, object> { ["ok"] = true, ["action"] = "read_range", ["values"] = val };
        }

        private object ClearRange(JObject a, string sheet, JObject pars)
        {
            Excel.Worksheet ws = GetSheet(sheet);
            string addr = (string)a["range"];
            string mode = (string)pars["mode"] ?? "contents";
            if (mode == "all") ws.Range[addr].Clear();
            else ws.Range[addr].ClearContents();
            return Ok("clear_range");
        }

        private object MergeCells(JObject a, string sheet)
        {
            GetSheet(sheet).Range[(string)a["range"]].Merge();
            return Ok("merge_cells");
        }

        private object UnmergeCells(JObject a, string sheet)
        {
            GetSheet(sheet).Range[(string)a["range"]].UnMerge();
            return Ok("unmerge_cells");
        }

        private object SetFormat(JObject a, string sheet, JObject pars)
        {
            Excel.Range rng = GetSheet(sheet).Range[(string)a["range"]];
            if (pars["bold"] != null) rng.Font.Bold = (bool)pars["bold"];
            if (pars["italic"] != null) rng.Font.Italic = (bool)pars["italic"];
            if (pars["font_size"] != null) rng.Font.Size = (double)pars["font_size"];
            if (pars["font_name"] != null) rng.Font.Name = (string)pars["font_name"];
            if (pars["font_color"] != null) rng.Font.Color = (int)pars["font_color"];
            if (pars["bg_color"] != null) rng.Interior.Color = (int)pars["bg_color"];
            if (pars["number_format"] != null) rng.NumberFormat = (string)pars["number_format"];
            return Ok("set_format");
        }

        private object InsertRows(JObject a, string sheet, JObject pars)
        {
            Excel.Worksheet ws = GetSheet(sheet);
            int row = (int)a["row"];
            int count = pars["count"]?.Value<int>() ?? 1;
            for (int i = 0; i < count; i++)
                ((Excel.Range)ws.Rows[row]).Insert();
            return Ok("insert_rows");
        }

        private object DeleteRows(JObject a, string sheet, JObject pars)
        {
            Excel.Worksheet ws = GetSheet(sheet);
            int row = (int)a["row"];
            int count = pars["count"]?.Value<int>() ?? 1;
            string addr = row + ":" + (row + count - 1);
            ((Excel.Range)ws.Rows[addr]).Delete();
            return Ok("delete_rows");
        }

        private object InsertCols(JObject a, string sheet, JObject pars)
        {
            Excel.Worksheet ws = GetSheet(sheet);
            int col = (int)a["col"];
            int count = pars["count"]?.Value<int>() ?? 1;
            for (int i = 0; i < count; i++)
                ((Excel.Range)ws.Columns[col]).Insert();
            return Ok("insert_cols");
        }

        private object DeleteCols(JObject a, string sheet, JObject pars)
        {
            Excel.Worksheet ws = GetSheet(sheet);
            int col = (int)a["col"];
            int count = pars["count"]?.Value<int>() ?? 1;
            string addr = col + ":" + (col + count - 1);
            ((Excel.Range)ws.Columns[addr]).Delete();
            return Ok("delete_cols");
        }

        private object AddSheet(JObject pars)
        {
            Excel.Workbook wb = Wb();
            Excel.Worksheet ws = (Excel.Worksheet)wb.Worksheets.Add();
            string name = (string)pars["name"];
            if (!string.IsNullOrEmpty(name)) ws.Name = name;
            return new Dictionary<string, object> { ["ok"] = true, ["action"] = "add_sheet", ["name"] = ws.Name };
        }

        private object RenameSheet(JObject a)
        {
            string oldName = (string)a["old_name"];
            string newName = (string)a["new_name"];
            ((Excel.Worksheet)Wb().Worksheets[oldName]).Name = newName;
            return Ok("rename_sheet");
        }

        private object DeleteSheet(string sheet)
        {
            _app.DisplayAlerts = false;
            GetSheet(sheet).Delete();
            _app.DisplayAlerts = true;
            return Ok("delete_sheet");
        }

        private object SortRange(JObject a, string sheet, JObject pars)
        {
            Excel.Worksheet ws = GetSheet(sheet);
            Excel.Range rng = ws.Range[(string)a["range"]];
            string keyCol = (string)a["key_col"];
            string order = (string)pars["order"] ?? "asc";
            Excel.XlSortOrder xlOrder = order == "desc" ? Excel.XlSortOrder.xlDescending : Excel.XlSortOrder.xlAscending;
            rng.Sort(Key1: ws.Range[keyCol], Order1: xlOrder, Header: Excel.XlYesNoGuess.xlYes);
            return Ok("sort_range");
        }

        private object FindReplace(JObject a, string sheet)
        {
            Excel.Worksheet ws = GetSheet(sheet);
            Excel.Range rng = ws.Range[(string)a["range"]];
            rng.Replace(What: (string)a["find"], Replacement: (string)a["replace"]);
            return Ok("find_replace");
        }

        private object ExportPdf(JObject a)
        {
            string path = (string)a["path"];
            Wb().ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, path);
            return Ok("export_pdf");
        }

        private static Dictionary<string, object> Ok(string action)
        {
            return new Dictionary<string, object> { ["ok"] = true, ["action"] = action };
        }
    }
}
