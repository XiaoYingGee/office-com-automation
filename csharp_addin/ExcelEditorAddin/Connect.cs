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
                case "autofit_columns": return AutofitColumns(action, sheet);
                case "add_sheet": return AddSheet(pars);
                case "rename_sheet": return RenameSheet(action);
                case "delete_sheet": return DeleteSheet(sheet);
                case "sort_range": return SortRange(action, sheet, pars);
                case "auto_filter": return AutoFilter(action, sheet, pars);
                case "find_replace": return FindReplace(action, sheet);
                case "calculate": _app.Calculate(); return Ok(name);
                case "export_pdf": return ExportPdf(action);
                case "copy_values": return CopyValues(action, sheet);
                case "paste_special": return PasteSpecial(action, sheet, pars);
                case "auto_fill": return AutoFill(action, sheet);
                case "set_border": return SetBorder(action, sheet, pars);
                case "add_conditional_format": return AddConditionalFormat(action, sheet, pars);
                case "clear_conditional_format": return ClearConditionalFormat(action, sheet);
                case "set_row_height": return SetRowHeight(action, sheet);
                case "set_col_width": return SetColWidth(action, sheet);
                case "group_rows": return GroupRows(action, sheet);
                case "ungroup_rows": return UngroupRows(action, sheet);
                case "group_cols": return GroupCols(action, sheet);
                case "ungroup_cols": return UngroupCols(action, sheet);
                case "hide_rows": return HideRows(action, sheet);
                case "unhide_rows": return UnhideRows(action, sheet);
                case "hide_cols": return HideCols(action, sheet);
                case "unhide_cols": return UnhideCols(action, sheet);
                case "copy_sheet": return CopySheet(sheet, pars);
                case "move_sheet": return MoveSheet(sheet, pars);
                case "protect_sheet": return ProtectSheet(sheet, pars);
                case "unprotect_sheet": return UnprotectSheet(sheet, pars);
                case "freeze_panes": return FreezePanes(action, sheet);
                case "unfreeze_panes": _app.ActiveWindow.FreezePanes = false; return Ok(name);
                case "add_validation": return AddValidation(action, sheet, pars);
                case "clear_validation": return ClearValidation(action, sheet);
                case "add_named_range": return AddNamedRange(action);
                case "delete_named_range": return DeleteNamedRange(action);
                case "add_chart": return AddChart(action, sheet, pars);
                case "delete_chart": return DeleteChart(action, sheet);
                case "modify_chart": return ModifyChart(action, sheet, pars);
                case "set_chart_title": return SetChartTitle(action, sheet);
                case "add_picture": return AddPicture(action, sheet, pars);
                case "delete_picture": return DeletePicture(action, sheet);
                case "add_comment": return AddComment(action, sheet, pars);
                case "delete_comment": return DeleteComment(action, sheet);
                case "add_hyperlink": return AddHyperlink(action, sheet, pars);
                case "delete_hyperlink": return DeleteHyperlink(action, sheet);
                case "set_page_setup": return SetPageSetup(sheet, pars);
                case "run_macro": return RunMacro(action);
                case "add_pivot_table": return AddPivotTable(action, sheet, pars);
                case "refresh_pivot": return RefreshPivot(action, sheet);
                case "export_image": return ExportImage(action, sheet, pars);
                case "inspect_sheet": return InspectSheetAction(action, sheet, pars);
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

        private object AutofitColumns(JObject a, string sheet)
        {
            Excel.Worksheet ws = GetSheet(sheet);
            string addr = (string)a["range"];
            if (!string.IsNullOrEmpty(addr))
                ws.Range[addr].Columns.AutoFit();
            else
                ws.UsedRange.Columns.AutoFit();
            return Ok("autofit_columns");
        }

        private object AutoFilter(JObject a, string sheet, JObject pars)
        {
            Excel.Worksheet ws = GetSheet(sheet);
            Excel.Range rng = ws.Range[(string)a["range"]];
            if (pars["field"] != null && pars["criteria"] != null)
                rng.AutoFilter((int)pars["field"], (string)pars["criteria"]);
            else
                rng.AutoFilter();
            return Ok("auto_filter");
        }

        private object CopyValues(JObject a, string sheet)
        {
            Excel.Worksheet ws = GetSheet(sheet);
            object val = ws.Range[(string)a["src_range"]].Value2;
            ws.Range[(string)a["dst_range"]].Value2 = val;
            return Ok("copy_values");
        }

        private object PasteSpecial(JObject a, string sheet, JObject pars)
        {
            Excel.Worksheet ws = GetSheet(sheet);
            ws.Range[(string)a["src_range"]].Copy();
            string pt = (string)pars["paste_type"] ?? "values";
            Excel.XlPasteType paste;
            switch (pt) {
                case "formulas": paste = Excel.XlPasteType.xlPasteFormulas; break;
                case "formats": paste = Excel.XlPasteType.xlPasteFormats; break;
                case "all": paste = Excel.XlPasteType.xlPasteAll; break;
                default: paste = Excel.XlPasteType.xlPasteValues; break;
            }
            ws.Range[(string)a["dst_range"]].PasteSpecial(paste);
            _app.CutCopyMode = Excel.XlCutCopyMode.xlCopy; // clear clipboard
            return Ok("paste_special");
        }

        private object AutoFill(JObject a, string sheet)
        {
            Excel.Worksheet ws = GetSheet(sheet);
            ws.Range[(string)a["src_range"]].AutoFill(ws.Range[(string)a["fill_range"]]);
            return Ok("auto_fill");
        }

        private object SetBorder(JObject a, string sheet, JObject pars)
        {
            Excel.Range rng = GetSheet(sheet).Range[(string)a["range"]];
            int style = pars["style"]?.Value<int>() ?? 1;
            int color = pars["color"]?.Value<int>() ?? 0;
            int weight = pars["weight"]?.Value<int>() ?? 2;
            string edges = (string)pars["edges"] ?? "all";
            int[] targets = edges == "all"
                ? new[] { 7, 8, 9, 10, 11, 12 }
                : ParseEdges(edges);
            foreach (int idx in targets)
            {
                Excel.Border b = rng.Borders[(Excel.XlBordersIndex)idx];
                b.LineStyle = style;
                b.Color = color;
                b.Weight = weight;
            }
            return Ok("set_border");
        }

        private static int[] ParseEdges(string edges)
        {
            var map = new Dictionary<string, int> {
                {"left",7},{"top",8},{"bottom",9},{"right",10},{"inside_h",12},{"inside_v",11}
            };
            var result = new List<int>();
            foreach (string e in edges.Split(','))
                if (map.TryGetValue(e.Trim(), out int v)) result.Add(v);
            return result.ToArray();
        }

        private object AddConditionalFormat(JObject a, string sheet, JObject pars)
        {
            Excel.Range rng = GetSheet(sheet).Range[(string)a["range"]];
            string rt = (string)pars["rule_type"] ?? "cell_value";
            int op = pars["operator"]?.Value<int>() ?? 3;
            string val = (string)pars["value"];
            Excel.FormatCondition cf;
            if (rt == "formula")
                cf = (Excel.FormatCondition)rng.FormatConditions.Add(Excel.XlFormatConditionType.xlExpression, Formula1: val);
            else
                cf = (Excel.FormatCondition)rng.FormatConditions.Add(Excel.XlFormatConditionType.xlCellValue, (Excel.XlFormatConditionOperator)op, val);
            JObject fmt = pars["format"] as JObject;
            if (fmt != null)
            {
                if (fmt["font_color"] != null) cf.Font.Color = (int)fmt["font_color"];
                if (fmt["bg_color"] != null) cf.Interior.Color = (int)fmt["bg_color"];
                if (fmt["bold"] != null) cf.Font.Bold = (bool)fmt["bold"];
            }
            return Ok("add_conditional_format");
        }

        private object ClearConditionalFormat(JObject a, string sheet)
        {
            GetSheet(sheet).Range[(string)a["range"]].FormatConditions.Delete();
            return Ok("clear_conditional_format");
        }

        private object SetRowHeight(JObject a, string sheet)
        {
            GetSheet(sheet).Rows[(int)a["row"]].RowHeight = (double)a["height"];
            return Ok("set_row_height");
        }

        private object SetColWidth(JObject a, string sheet)
        {
            GetSheet(sheet).Columns[(int)a["col"]].ColumnWidth = (double)a["width"];
            return Ok("set_col_width");
        }

        private object GroupRows(JObject a, string sheet)
        {
            int s = (int)a["start_row"], e = (int)a["end_row"];
            ((Excel.Range)GetSheet(sheet).Rows[s + ":" + e]).Group();
            return Ok("group_rows");
        }

        private object UngroupRows(JObject a, string sheet)
        {
            int s = (int)a["start_row"], e = (int)a["end_row"];
            ((Excel.Range)GetSheet(sheet).Rows[s + ":" + e]).Ungroup();
            return Ok("ungroup_rows");
        }

        private object GroupCols(JObject a, string sheet)
        {
            string s = (string)a["start_col"], e = (string)a["end_col"];
            ((Excel.Range)GetSheet(sheet).Columns[s + ":" + e]).Group();
            return Ok("group_cols");
        }

        private object UngroupCols(JObject a, string sheet)
        {
            string s = (string)a["start_col"], e = (string)a["end_col"];
            ((Excel.Range)GetSheet(sheet).Columns[s + ":" + e]).Ungroup();
            return Ok("ungroup_cols");
        }

        private object HideRows(JObject a, string sheet)
        {
            int s = (int)a["start_row"], e = (int)a["end_row"];
            ((Excel.Range)GetSheet(sheet).Rows[s + ":" + e]).Hidden = true;
            return Ok("hide_rows");
        }

        private object UnhideRows(JObject a, string sheet)
        {
            int s = (int)a["start_row"], e = (int)a["end_row"];
            ((Excel.Range)GetSheet(sheet).Rows[s + ":" + e]).Hidden = false;
            return Ok("unhide_rows");
        }

        private object HideCols(JObject a, string sheet)
        {
            string s = (string)a["start_col"], e = (string)a["end_col"];
            ((Excel.Range)GetSheet(sheet).Columns[s + ":" + e]).Hidden = true;
            return Ok("hide_cols");
        }

        private object UnhideCols(JObject a, string sheet)
        {
            string s = (string)a["start_col"], e = (string)a["end_col"];
            ((Excel.Range)GetSheet(sheet).Columns[s + ":" + e]).Hidden = false;
            return Ok("unhide_cols");
        }

        private object CopySheet(string sheet, JObject pars)
        {
            Excel.Worksheet ws = GetSheet(sheet);
            if (pars["after"] != null)
                ws.Copy(After: GetSheet(pars["after"].ToString()));
            else if (pars["before"] != null)
                ws.Copy(Before: GetSheet(pars["before"].ToString()));
            else
                ws.Copy(After: (Excel.Worksheet)Wb().Worksheets[Wb().Worksheets.Count]);
            return new Dictionary<string, object> { ["ok"] = true, ["action"] = "copy_sheet", ["name"] = _app.ActiveSheet.Name };
        }

        private object MoveSheet(string sheet, JObject pars)
        {
            Excel.Worksheet ws = GetSheet(sheet);
            if (pars["after"] != null)
                ws.Move(After: GetSheet(pars["after"].ToString()));
            else if (pars["before"] != null)
                ws.Move(Before: GetSheet(pars["before"].ToString()));
            return Ok("move_sheet");
        }

        private object ProtectSheet(string sheet, JObject pars)
        {
            Excel.Worksheet ws = GetSheet(sheet);
            string pw = (string)pars["password"];
            if (!string.IsNullOrEmpty(pw)) ws.Protect(pw);
            else ws.Protect();
            return Ok("protect_sheet");
        }

        private object UnprotectSheet(string sheet, JObject pars)
        {
            Excel.Worksheet ws = GetSheet(sheet);
            string pw = (string)pars["password"];
            if (!string.IsNullOrEmpty(pw)) ws.Unprotect(pw);
            else ws.Unprotect();
            return Ok("unprotect_sheet");
        }

        private object FreezePanes(JObject a, string sheet)
        {
            Excel.Worksheet ws = GetSheet(sheet);
            int row = (int)a["row"], col = (int)a["col"];
            _app.Goto(ws.Cells[row, col]);
            _app.ActiveWindow.FreezePanes = true;
            return Ok("freeze_panes");
        }

        private object AddValidation(JObject a, string sheet, JObject pars)
        {
            Excel.Range rng = GetSheet(sheet).Range[(string)a["range"]];
            rng.Validation.Delete();
            string vt = (string)pars["type"] ?? "list";
            int xlType = vt == "whole" ? 1 : vt == "decimal" ? 2 : vt == "text_length" ? 6 : 3;
            string formula = (string)pars["formula"];
            JArray vals = pars["values"] as JArray;
            if (vt == "list" && vals != null)
                formula = string.Join(",", vals);
            if (!string.IsNullOrEmpty(formula))
                rng.Validation.Add((Excel.XlDVType)xlType, Excel.XlDVAlertStyle.xlValidAlertStop,
                    Excel.XlFormatConditionOperator.xlBetween, formula);
            return Ok("add_validation");
        }

        private object ClearValidation(JObject a, string sheet)
        {
            GetSheet(sheet).Range[(string)a["range"]].Validation.Delete();
            return Ok("clear_validation");
        }

        private object AddNamedRange(JObject a)
        {
            Wb().Names.Add((string)a["name"], RefersTo: (string)a["refers_to"]);
            return Ok("add_named_range");
        }

        private object DeleteNamedRange(JObject a)
        {
            Wb().Names.Item((string)a["name"]).Delete();
            return Ok("delete_named_range");
        }

        private object AddChart(JObject a, string sheet, JObject pars)
        {
            Excel.Worksheet ws = GetSheet(sheet);
            double left = pars["left"]?.Value<double>() ?? 100;
            double top = pars["top"]?.Value<double>() ?? 100;
            double width = pars["width"]?.Value<double>() ?? 400;
            double height = pars["height"]?.Value<double>() ?? 300;
            Excel.ChartObject co = ((Excel.ChartObjects)ws.ChartObjects()).Add(left, top, width, height);
            string ct = (string)pars["chart_type"] ?? "xlColumnClustered";
            co.Chart.ChartType = MapChartType(ct);
            string dr = (string)a["data_range"];
            if (!string.IsNullOrEmpty(dr))
                co.Chart.SetSourceData(ws.Range[dr]);
            return new Dictionary<string, object> { ["ok"] = true, ["action"] = "add_chart", ["name"] = co.Name };
        }

        private object DeleteChart(JObject a, string sheet)
        {
            ((Excel.ChartObjects)GetSheet(sheet).ChartObjects())[(string)a["chart_name"]].Delete();
            return Ok("delete_chart");
        }

        private object ModifyChart(JObject a, string sheet, JObject pars)
        {
            string cn = (string)a["chart_name"];
            string ct = (string)pars["chart_type"] ?? "xlColumnClustered";
            ((Excel.ChartObjects)GetSheet(sheet).ChartObjects())[cn].Chart.ChartType = MapChartType(ct);
            return Ok("modify_chart");
        }

        private object SetChartTitle(JObject a, string sheet)
        {
            string cn = (string)a["chart_name"];
            Excel.Chart chart = ((Excel.ChartObjects)GetSheet(sheet).ChartObjects())[cn].Chart;
            chart.HasTitle = true;
            chart.ChartTitle.Text = (string)a["title"];
            return Ok("set_chart_title");
        }

        private static Excel.XlChartType MapChartType(string ct)
        {
            switch (ct)
            {
                case "xlLine": return (Excel.XlChartType)4;
                case "xlPie": return (Excel.XlChartType)5;
                case "xlBarClustered": return (Excel.XlChartType)57;
                case "xlArea": return (Excel.XlChartType)1;
                case "xlXYScatter": return (Excel.XlChartType)(-4169);
                default: return (Excel.XlChartType)51; // xlColumnClustered
            }
        }

        private object AddPicture(JObject a, string sheet, JObject pars)
        {
            Excel.Worksheet ws = GetSheet(sheet);
            string path = (string)a["path"];
            float left = (float)(pars["left"]?.Value<double>() ?? 0);
            float top = (float)(pars["top"]?.Value<double>() ?? 0);
            float width = (float)(pars["width"]?.Value<double>() ?? -1);
            float height = (float)(pars["height"]?.Value<double>() ?? -1);
            Excel.Shape pic = ws.Shapes.AddPicture(path, Office.MsoTriState.msoFalse,
                Office.MsoTriState.msoTrue, left, top, width > 0 ? width : -1, height > 0 ? height : -1);
            return new Dictionary<string, object> { ["ok"] = true, ["action"] = "add_picture", ["name"] = pic.Name };
        }

        private object DeletePicture(JObject a, string sheet)
        {
            GetSheet(sheet).Shapes.Item((string)a["name"]).Delete();
            return Ok("delete_picture");
        }

        private object AddComment(JObject a, string sheet, JObject pars)
        {
            Excel.Range rng = GetSheet(sheet).Range[(string)a["cell"]];
            try { rng.Comment.Delete(); } catch { }
            rng.AddComment((string)a["text"]);
            return Ok("add_comment");
        }

        private object DeleteComment(JObject a, string sheet)
        {
            GetSheet(sheet).Range[(string)a["cell"]].Comment.Delete();
            return Ok("delete_comment");
        }

        private object AddHyperlink(JObject a, string sheet, JObject pars)
        {
            Excel.Worksheet ws = GetSheet(sheet);
            Excel.Range rng = ws.Range[(string)a["cell"]];
            string url = (string)a["url"];
            string display = (string)pars["display_text"] ?? url;
            ws.Hyperlinks.Add(rng, url, TextToDisplay: display);
            return Ok("add_hyperlink");
        }

        private object DeleteHyperlink(JObject a, string sheet)
        {
            Excel.Range rng = GetSheet(sheet).Range[(string)a["cell"]];
            foreach (Excel.Hyperlink hl in GetSheet(sheet).Hyperlinks)
                if (hl.Range.Address == rng.Address) { hl.Delete(); break; }
            return Ok("delete_hyperlink");
        }

        private object SetPageSetup(string sheet, JObject pars)
        {
            Excel.PageSetup ps = GetSheet(sheet).PageSetup;
            if (pars["orientation"] != null)
                ps.Orientation = (string)pars["orientation"] == "landscape"
                    ? Excel.XlPageOrientation.xlLandscape : Excel.XlPageOrientation.xlPortrait;
            if (pars["paper_size"] != null) ps.PaperSize = (Excel.XlPaperSize)(int)pars["paper_size"];
            if (pars["left_margin"] != null) ps.LeftMargin = (double)pars["left_margin"];
            if (pars["right_margin"] != null) ps.RightMargin = (double)pars["right_margin"];
            if (pars["top_margin"] != null) ps.TopMargin = (double)pars["top_margin"];
            if (pars["bottom_margin"] != null) ps.BottomMargin = (double)pars["bottom_margin"];
            if (pars["header"] != null) ps.CenterHeader = (string)pars["header"];
            if (pars["footer"] != null) ps.CenterFooter = (string)pars["footer"];
            return Ok("set_page_setup");
        }

        private object RunMacro(JObject a)
        {
            string macro = (string)a["macro_name"];
            JArray args = a["args"] as JArray;
            object result;
            if (args != null && args.Count > 0)
                result = _app.Run(macro, args[0]?.ToObject<object>());
            else
                result = _app.Run(macro);
            return new Dictionary<string, object> { ["ok"] = true, ["action"] = "run_macro", ["result"] = result?.ToString() };
        }

        private object AddPivotTable(JObject a, string sheet, JObject pars)
        {
            Excel.Worksheet ws = GetSheet(sheet);
            Excel.Range src = ws.Range[(string)a["source_range"]];
            Excel.PivotCache pc = Wb().PivotCaches().Create(Excel.XlPivotTableSourceType.xlDatabase, src);
            Excel.Worksheet destWs = GetSheet(a["dest_sheet"]?.ToString() ?? sheet);
            string tblName = (string)pars["table_name"] ?? "PivotTable1";
            Excel.PivotTable pt = pc.CreatePivotTable(destWs.Range[(string)a["dest_cell"]], tblName);
            JArray rows = pars["rows"] as JArray;
            if (rows != null)
                foreach (JToken r in rows) pt.PivotFields(r.ToString()).Orientation = Excel.XlPivotFieldOrientation.xlRowField;
            JArray cols = pars["cols"] as JArray;
            if (cols != null)
                foreach (JToken c in cols) pt.PivotFields(c.ToString()).Orientation = Excel.XlPivotFieldOrientation.xlColumnField;
            JArray vals = pars["values"] as JArray;
            if (vals != null)
                foreach (JToken v in vals) pt.AddDataField(pt.PivotFields(v.ToString()));
            return new Dictionary<string, object> { ["ok"] = true, ["action"] = "add_pivot_table", ["name"] = tblName };
        }

        private object RefreshPivot(JObject a, string sheet)
        {
            GetSheet(sheet).PivotTables((string)a["table_name"]).RefreshTable();
            return Ok("refresh_pivot");
        }

        private object ExportImage(JObject a, string sheet, JObject pars)
        {
            Excel.Worksheet ws = GetSheet(sheet);
            ws.UsedRange.CopyPicture(Excel.XlPictureAppearance.xlScreen, Excel.XlCopyPictureFormat.xlBitmap);
            Excel.ChartObject co = ((Excel.ChartObjects)ws.ChartObjects()).Add(0, 0,
                pars["width"]?.Value<int>() ?? 1024, pars["height"]?.Value<int>() ?? 768);
            co.Chart.Paste();
            co.Chart.Export((string)a["path"], "PNG");
            co.Delete();
            return Ok("export_image");
        }

        private object InspectSheetAction(JObject a, string sheet, JObject pars)
        {
            Excel.Worksheet ws = GetSheet(sheet);
            Excel.Range used = ws.UsedRange;
            int maxRows = pars["max_rows"]?.Value<int>() ?? 50;
            int maxCols = pars["max_cols"]?.Value<int>() ?? 26;
            int rows = Math.Min(used?.Rows.Count ?? 0, maxRows);
            int cols = Math.Min(used?.Columns.Count ?? 0, maxCols);
            var data = new List<List<object>>();
            int startRow = used?.Row ?? 1;
            int startCol = used?.Column ?? 1;
            for (int r = startRow; r < startRow + rows; r++)
            {
                var row = new List<object>();
                for (int c = startCol; c < startCol + cols; c++)
                    row.Add(((Excel.Range)ws.Cells[r, c]).Value2);
                data.Add(row);
            }
            return new Dictionary<string, object> { ["ok"] = true, ["action"] = "inspect_sheet",
                ["name"] = ws.Name, ["used_range"] = used?.Address ?? "", ["data"] = data };
        }

        private static Dictionary<string, object> Ok(string action)
        {
            return new Dictionary<string, object> { ["ok"] = true, ["action"] = action };
        }
    }
}
