"""
Excel COM 自动化引擎 — 支持 pywin32 / VBA 两种 backend

用法:
  python excel_editor.py <xlsx文件> --inspect
  python excel_editor.py <xlsx文件> --exec-actions '[{"action":"write_cell","sheet":"Sheet1","cell":"A1","value":42}]'
  python excel_editor.py <xlsx文件> --exec-actions actions.json
  python excel_editor.py <xlsx文件> --exec-script edit.py
  python excel_editor.py --create <xlsx文件>
  python excel_editor.py <xlsx文件> --interactive-actions
"""
import sys
import os
import json
import argparse
import time

try:
    import win32com.client
    import pythoncom
except ImportError:
    print("pip install pywin32 (Windows + Office required)")
    sys.exit(1)

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass


class ExcelCOM:
    """Excel COM automation engine — pywin32 backend (baseline)."""

    def __init__(self, visible=False):
        pythoncom.CoInitialize()
        try:
            self.app = win32com.client.gencache.EnsureDispatch("Excel.Application")
        except Exception:
            self.app = win32com.client.Dispatch("Excel.Application")
        self.app.Visible = visible
        self.app.DisplayAlerts = False
        self.wb = None
        self.filepath = None

    def open(self, path):
        self.filepath = os.path.abspath(path)
        self.wb = self.app.Workbooks.Open(self.filepath)
        print(f"Opened: {self.filepath} ({self.wb.Worksheets.Count} sheets)")
        return self

    def create(self, path):
        self.filepath = os.path.abspath(path)
        self.wb = self.app.Workbooks.Add()
        self.wb.SaveAs(self.filepath, 51)  # xlOpenXMLWorkbook
        print(f"Created: {self.filepath}")
        return self

    def save(self, path=None):
        if path:
            self.wb.SaveAs(os.path.abspath(path), 51)
        else:
            self.wb.Save()

    def close(self):
        try:
            if self.wb:
                self.wb.Close(False)
            if self.app:
                self.app.Quit()
        except Exception:
            pass
        pythoncom.CoUninitialize()

    # ---- Inspect ----
    def inspect(self):
        result = {"sheets": []}
        for si in range(1, self.wb.Worksheets.Count + 1):
            ws = self.wb.Worksheets(si)
            used = ws.UsedRange
            sheet_info = {
                "index": si,
                "name": ws.Name,
                "used_range": used.Address if used else "",
                "rows": used.Rows.Count if used else 0,
                "cols": used.Columns.Count if used else 0,
            }
            result["sheets"].append(sheet_info)
        return result

    def inspect_sheet(self, sheet_name_or_index, max_rows=50, max_cols=26):
        ws = self._get_sheet(sheet_name_or_index)
        used = ws.UsedRange
        if not used:
            return {"name": ws.Name, "data": []}

        rows = min(used.Rows.Count, max_rows)
        cols = min(used.Columns.Count, max_cols)
        start_row = used.Row
        start_col = used.Column

        data = []
        for r in range(start_row, start_row + rows):
            row_data = []
            for c in range(start_col, start_col + cols):
                val = ws.Cells(r, c).Value2
                row_data.append(val)
            data.append(row_data)

        return {
            "name": ws.Name,
            "used_range": used.Address,
            "data": data,
        }

    # ---- Cell I/O ----
    def read_cell(self, sheet, cell, property="value2"):
        ws = self._get_sheet(sheet)
        rng = ws.Range(cell)
        if property == "formula":
            return rng.Formula
        elif property == "text":
            return rng.Text
        else:
            return rng.Value2

    def write_cell(self, sheet, cell, value, kind="auto"):
        ws = self._get_sheet(sheet)
        rng = ws.Range(cell)
        if kind == "formula":
            rng.Formula = value
        else:
            rng.Value2 = value

    def write_range(self, sheet, range_addr, values):
        ws = self._get_sheet(sheet)
        rng = ws.Range(range_addr)
        rng.Value2 = values

    def read_range(self, sheet, range_addr):
        ws = self._get_sheet(sheet)
        rng = ws.Range(range_addr)
        val = rng.Value2
        if val is None:
            return [[None]]
        if not isinstance(val, tuple):
            return [[val]]
        return [list(row) for row in val]

    def clear_range(self, sheet, range_addr, mode="contents"):
        ws = self._get_sheet(sheet)
        rng = ws.Range(range_addr)
        if mode == "all":
            rng.Clear()
        else:
            rng.ClearContents()

    def merge_cells(self, sheet, range_addr):
        ws = self._get_sheet(sheet)
        ws.Range(range_addr).Merge()

    def unmerge_cells(self, sheet, range_addr):
        ws = self._get_sheet(sheet)
        ws.Range(range_addr).UnMerge()

    # ---- Formatting ----
    def set_format(self, sheet, range_addr, **kwargs):
        ws = self._get_sheet(sheet)
        rng = ws.Range(range_addr)
        if "bold" in kwargs:
            rng.Font.Bold = kwargs["bold"]
        if "italic" in kwargs:
            rng.Font.Italic = kwargs["italic"]
        if "font_size" in kwargs:
            rng.Font.Size = kwargs["font_size"]
        if "font_name" in kwargs:
            rng.Font.Name = kwargs["font_name"]
        if "font_color" in kwargs:
            rng.Font.Color = kwargs["font_color"]
        if "bg_color" in kwargs:
            rng.Interior.Color = kwargs["bg_color"]
        if "number_format" in kwargs:
            rng.NumberFormat = kwargs["number_format"]
        if "h_align" in kwargs:
            rng.HorizontalAlignment = kwargs["h_align"]
        if "v_align" in kwargs:
            rng.VerticalAlignment = kwargs["v_align"]
        if "merge" in kwargs and kwargs["merge"]:
            rng.Merge()
        if "wrap_text" in kwargs:
            rng.WrapText = kwargs["wrap_text"]

    # ---- Row/Column ----
    def insert_rows(self, sheet, row, count=1):
        ws = self._get_sheet(sheet)
        for _ in range(count):
            ws.Rows(row).Insert()

    def delete_rows(self, sheet, row, count=1):
        ws = self._get_sheet(sheet)
        ws.Rows(f"{row}:{row + count - 1}").Delete()

    def insert_cols(self, sheet, col, count=1):
        ws = self._get_sheet(sheet)
        for _ in range(count):
            ws.Columns(col).Insert()

    def delete_cols(self, sheet, col, count=1):
        ws = self._get_sheet(sheet)
        ws.Columns(f"{col}:{col + count - 1}").Delete()

    def autofit_columns(self, sheet, range_addr=None):
        ws = self._get_sheet(sheet)
        if range_addr:
            ws.Range(range_addr).Columns.AutoFit()
        else:
            ws.UsedRange.Columns.AutoFit()

    # ---- Worksheet ----
    def add_sheet(self, name=None, after=None):
        if after:
            ws = self.wb.Worksheets.Add(After=self._get_sheet(after))
        else:
            ws = self.wb.Worksheets.Add()
        if name:
            ws.Name = name
        return ws.Name

    def rename_sheet(self, old_name, new_name):
        ws = self._get_sheet(old_name)
        ws.Name = new_name

    def delete_sheet(self, sheet):
        ws = self._get_sheet(sheet)
        ws.Delete()

    # ---- Formula ----
    def calculate(self):
        self.app.Calculate()

    # ---- Sort / Filter ----
    def sort_range(self, sheet, range_addr, key_col, order="asc"):
        ws = self._get_sheet(sheet)
        rng = ws.Range(range_addr)
        xl_asc = 1
        xl_desc = 2
        rng.Sort(
            Key1=ws.Range(key_col),
            Order1=xl_asc if order == "asc" else xl_desc,
            Header=1,  # xlYes
        )

    def auto_filter(self, sheet, range_addr, field=None, criteria=None):
        ws = self._get_sheet(sheet)
        rng = ws.Range(range_addr)
        if field and criteria:
            rng.AutoFilter(Field=field, Criteria1=criteria)
        else:
            rng.AutoFilter()

    def find_replace(self, sheet, range_addr, find_text, replace_text):
        ws = self._get_sheet(sheet)
        rng = ws.Range(range_addr)
        rng.Replace(What=find_text, Replacement=replace_text)

    # ---- Export ----
    def export_pdf(self, output_path):
        self.wb.ExportAsFixedFormat(0, os.path.abspath(output_path))

    def export_image(self, sheet, output_path, width=1024, height=768):
        ws = self._get_sheet(sheet)
        chart = self.wb.Charts.Add()
        chart.Location(2, ws.Name)  # xlLocationAsObject
        ws.ChartObjects(ws.ChartObjects().Count).Delete()
        ws.UsedRange.CopyPicture(1, 2)  # xlScreen, xlBitmap
        chart_obj = ws.ChartObjects().Add(0, 0, width, height)
        chart_obj.Chart.Paste()
        chart_obj.Chart.Export(os.path.abspath(output_path), "PNG")
        chart_obj.Delete()

    # ---- Area operations ----
    def copy_values(self, sheet, src_range, dst_range):
        ws = self._get_sheet(sheet)
        val = ws.Range(src_range).Value2
        ws.Range(dst_range).Value2 = val

    def paste_special(self, sheet, src_range, dst_range, paste_type="values"):
        ws = self._get_sheet(sheet)
        ws.Range(src_range).Copy()
        paste_map = {"values": -4163, "formulas": -4123, "formats": -4122, "all": -4104}
        ws.Range(dst_range).PasteSpecial(Paste=paste_map.get(paste_type, -4163))
        self.app.CutCopyMode = False

    def auto_fill(self, sheet, src_range, fill_range):
        ws = self._get_sheet(sheet)
        ws.Range(src_range).AutoFill(ws.Range(fill_range))

    # ---- Borders ----
    def set_border(self, sheet, range_addr, style=1, color=0, weight=2, edges="all"):
        ws = self._get_sheet(sheet)
        rng = ws.Range(range_addr)
        edge_map = {"left": 7, "right": 10, "top": 8, "bottom": 9,
                    "inside_h": 12, "inside_v": 11}
        if edges == "all":
            targets = [7, 8, 9, 10, 11, 12]
        else:
            targets = [edge_map[e.strip()] for e in edges.split(",") if e.strip() in edge_map]
        for idx in targets:
            border = rng.Borders(idx)
            border.LineStyle = style
            border.Color = color
            border.Weight = weight

    # ---- Conditional formatting ----
    def add_conditional_format(self, sheet, range_addr, rule_type="cell_value",
                               operator=3, value=None, format_params=None):
        ws = self._get_sheet(sheet)
        rng = ws.Range(range_addr)
        if rule_type == "cell_value":
            cf = rng.FormatConditions.Add(1, operator, value)
        elif rule_type == "formula":
            cf = rng.FormatConditions.Add(2, Formula1=value)
        else:
            return
        if format_params:
            if "font_color" in format_params:
                cf.Font.Color = format_params["font_color"]
            if "bg_color" in format_params:
                cf.Interior.Color = format_params["bg_color"]
            if "bold" in format_params:
                cf.Font.Bold = format_params["bold"]

    def clear_conditional_format(self, sheet, range_addr):
        ws = self._get_sheet(sheet)
        ws.Range(range_addr).FormatConditions.Delete()

    # ---- Row/Col dimensions ----
    def set_row_height(self, sheet, row, height):
        ws = self._get_sheet(sheet)
        ws.Rows(row).RowHeight = height

    def set_col_width(self, sheet, col, width):
        ws = self._get_sheet(sheet)
        ws.Columns(col).ColumnWidth = width

    # ---- Grouping ----
    def group_rows(self, sheet, start_row, end_row):
        ws = self._get_sheet(sheet)
        ws.Rows(f"{start_row}:{end_row}").Group()

    def ungroup_rows(self, sheet, start_row, end_row):
        ws = self._get_sheet(sheet)
        ws.Rows(f"{start_row}:{end_row}").Ungroup()

    def group_cols(self, sheet, start_col, end_col):
        ws = self._get_sheet(sheet)
        ws.Columns(f"{start_col}:{end_col}").Group()

    def ungroup_cols(self, sheet, start_col, end_col):
        ws = self._get_sheet(sheet)
        ws.Columns(f"{start_col}:{end_col}").Ungroup()

    # ---- Hide/Unhide ----
    def hide_rows(self, sheet, start_row, end_row):
        ws = self._get_sheet(sheet)
        ws.Rows(f"{start_row}:{end_row}").Hidden = True

    def unhide_rows(self, sheet, start_row, end_row):
        ws = self._get_sheet(sheet)
        ws.Rows(f"{start_row}:{end_row}").Hidden = False

    def hide_cols(self, sheet, start_col, end_col):
        ws = self._get_sheet(sheet)
        ws.Columns(f"{start_col}:{end_col}").Hidden = True

    def unhide_cols(self, sheet, start_col, end_col):
        ws = self._get_sheet(sheet)
        ws.Columns(f"{start_col}:{end_col}").Hidden = False

    # ---- Sheet management ----
    def copy_sheet(self, sheet, before=None, after=None):
        ws = self._get_sheet(sheet)
        if after:
            ws.Copy(After=self._get_sheet(after))
        elif before:
            ws.Copy(Before=self._get_sheet(before))
        else:
            ws.Copy(After=self.wb.Worksheets(self.wb.Worksheets.Count))
        return self.app.ActiveSheet.Name

    def move_sheet(self, sheet, before=None, after=None):
        ws = self._get_sheet(sheet)
        if after:
            ws.Move(After=self._get_sheet(after))
        elif before:
            ws.Move(Before=self._get_sheet(before))

    def protect_sheet(self, sheet, password=None):
        ws = self._get_sheet(sheet)
        if password:
            ws.Protect(Password=password)
        else:
            ws.Protect()

    def unprotect_sheet(self, sheet, password=None):
        ws = self._get_sheet(sheet)
        if password:
            ws.Unprotect(Password=password)
        else:
            ws.Unprotect()

    def freeze_panes(self, sheet, row, col):
        ws = self._get_sheet(sheet)
        self.app.Goto(ws.Cells(row, col))
        self.app.ActiveWindow.FreezePanes = True

    def unfreeze_panes(self):
        self.app.ActiveWindow.FreezePanes = False

    # ---- Data operations ----
    def add_validation(self, sheet, range_addr, val_type="list", formula=None, values=None):
        ws = self._get_sheet(sheet)
        rng = ws.Range(range_addr)
        rng.Validation.Delete()
        type_map = {"list": 3, "whole": 1, "decimal": 2, "text_length": 6}
        xl_type = type_map.get(val_type, 3)
        if val_type == "list" and values:
            formula1 = ",".join(str(v) for v in values)
            rng.Validation.Add(xl_type, 1, 1, formula1)
        elif formula:
            rng.Validation.Add(xl_type, 1, 1, formula)

    def clear_validation(self, sheet, range_addr):
        ws = self._get_sheet(sheet)
        ws.Range(range_addr).Validation.Delete()

    def add_named_range(self, name, refers_to):
        self.wb.Names.Add(Name=name, RefersTo=refers_to)

    def delete_named_range(self, name):
        self.wb.Names(name).Delete()

    # ---- Charts ----
    def add_chart(self, sheet, chart_type="xlColumnClustered", data_range=None,
                  left=100, top=100, width=400, height=300):
        ws = self._get_sheet(sheet)
        chart_obj = ws.ChartObjects().Add(left, top, width, height)
        chart = chart_obj.Chart
        type_map = {"xlColumnClustered": 51, "xlLine": 4, "xlPie": 5,
                    "xlBarClustered": 57, "xlArea": 1, "xlXYScatter": -4169}
        chart.ChartType = type_map.get(chart_type, 51)
        if data_range:
            chart.SetSourceData(ws.Range(data_range))
        return chart_obj.Name

    def delete_chart(self, sheet, chart_name):
        ws = self._get_sheet(sheet)
        ws.ChartObjects(chart_name).Delete()

    def modify_chart(self, sheet, chart_name, chart_type):
        ws = self._get_sheet(sheet)
        type_map = {"xlColumnClustered": 51, "xlLine": 4, "xlPie": 5,
                    "xlBarClustered": 57, "xlArea": 1, "xlXYScatter": -4169}
        ws.ChartObjects(chart_name).Chart.ChartType = type_map.get(chart_type, 51)

    def set_chart_title(self, sheet, chart_name, title):
        ws = self._get_sheet(sheet)
        chart = ws.ChartObjects(chart_name).Chart
        chart.HasTitle = True
        chart.ChartTitle.Text = title

    # ---- Pictures ----
    def add_picture(self, sheet, path, left=0, top=0, width=-1, height=-1):
        ws = self._get_sheet(sheet)
        pic = ws.Shapes.AddPicture(
            os.path.abspath(path), False, True, left, top,
            width if width > 0 else -1, height if height > 0 else -1)
        return pic.Name

    def delete_picture(self, sheet, name):
        ws = self._get_sheet(sheet)
        ws.Shapes(name).Delete()

    # ---- Comments ----
    def add_comment(self, sheet, cell, text, author=None):
        ws = self._get_sheet(sheet)
        rng = ws.Range(cell)
        try:
            rng.Comment.Delete()
        except Exception:
            pass
        rng.AddComment(text)
        if author:
            rng.Comment.Author = author

    def delete_comment(self, sheet, cell):
        ws = self._get_sheet(sheet)
        ws.Range(cell).Comment.Delete()

    # ---- Hyperlinks ----
    def add_hyperlink(self, sheet, cell, url, display_text=None):
        ws = self._get_sheet(sheet)
        rng = ws.Range(cell)
        ws.Hyperlinks.Add(Anchor=rng, Address=url,
                          TextToDisplay=display_text or url)

    def delete_hyperlink(self, sheet, cell):
        ws = self._get_sheet(sheet)
        rng = ws.Range(cell)
        for hl in ws.Hyperlinks:
            if hl.Range.Address == rng.Address:
                hl.Delete()
                break

    # ---- Page setup ----
    def set_page_setup(self, sheet, **kwargs):
        ws = self._get_sheet(sheet)
        ps = ws.PageSetup
        if "orientation" in kwargs:
            ps.Orientation = 2 if kwargs["orientation"] == "landscape" else 1
        if "paper_size" in kwargs:
            ps.PaperSize = kwargs["paper_size"]
        if "left_margin" in kwargs:
            ps.LeftMargin = kwargs["left_margin"]
        if "right_margin" in kwargs:
            ps.RightMargin = kwargs["right_margin"]
        if "top_margin" in kwargs:
            ps.TopMargin = kwargs["top_margin"]
        if "bottom_margin" in kwargs:
            ps.BottomMargin = kwargs["bottom_margin"]
        if "header" in kwargs:
            ps.CenterHeader = kwargs["header"]
        if "footer" in kwargs:
            ps.CenterFooter = kwargs["footer"]
        if "fit_to_pages_wide" in kwargs:
            ps.FitToPagesWide = kwargs["fit_to_pages_wide"]
        if "fit_to_pages_tall" in kwargs:
            ps.FitToPagesTall = kwargs["fit_to_pages_tall"]

    # ---- Macros ----
    def run_macro(self, macro_name, *args):
        if args:
            return self.app.Run(macro_name, *args)
        return self.app.Run(macro_name)

    # ---- Pivot tables ----
    def add_pivot_table(self, sheet, source_range, dest_sheet, dest_cell,
                        table_name="PivotTable1", rows=None, cols=None, values=None):
        ws = self._get_sheet(sheet)
        src = ws.Range(source_range)
        pc = self.wb.PivotCaches().Create(1, src)  # xlDatabase
        dest_ws = self._get_sheet(dest_sheet)
        pt = pc.CreatePivotTable(dest_ws.Range(dest_cell), table_name)
        if rows:
            for field_name in rows:
                pt.PivotFields(field_name).Orientation = 1  # xlRowField
        if cols:
            for field_name in cols:
                pt.PivotFields(field_name).Orientation = 2  # xlColumnField
        if values:
            for field_name in values:
                pt.AddDataField(pt.PivotFields(field_name))
        return table_name

    def refresh_pivot(self, sheet, table_name):
        ws = self._get_sheet(sheet)
        ws.PivotTables(table_name).RefreshTable()

    # ---- Helpers ----
    def _get_sheet(self, name_or_index):
        if isinstance(name_or_index, int):
            return self.wb.Worksheets(name_or_index)
        return self.wb.Worksheets(name_or_index)


class ExcelVBA(ExcelCOM):
    """VBA backend — delegates heavy ops to in-process VBA macros via Application.Run."""

    def __init__(self, visible=False, macro_module="ExcelEditorBridge"):
        super().__init__(visible=visible)
        self.macro_module = macro_module
        self._vba_injected = False

    def open(self, path):
        super().open(path)
        self._inject_vba()
        return self

    def create(self, path):
        super().create(path)
        self._inject_vba()
        return self

    def _inject_vba(self):
        if self._vba_injected:
            return
        bas_path = os.path.join(
            os.path.dirname(os.path.abspath(__file__)),
            "..", "references", "ExcelEditorBridge.bas"
        )
        bas_path = os.path.normpath(bas_path)
        if not os.path.isfile(bas_path):
            raise FileNotFoundError(f"VBA module not found: {bas_path}")
        try:
            vb_proj = self.wb.VBProject
            vb_proj.VBComponents.Import(bas_path)
            self._vba_injected = True
        except Exception as e:
            raise RuntimeError(
                f"Failed to inject VBA module. Enable 'Trust access to VBA project object model' "
                f"in Excel Trust Center. Error: {e}"
            )

    def inspect(self):
        raw = self.app.Run(f"{self.macro_module}.InspectWorkbookJson")
        return json.loads(raw)

    def inspect_sheet(self, sheet_name_or_index, max_rows=50, max_cols=26):
        params = json.dumps({"sheet": sheet_name_or_index, "max_rows": max_rows, "max_cols": max_cols})
        raw = self.app.Run(f"{self.macro_module}.InspectSheetJson", params)
        return json.loads(raw)

    def execute_actions(self, actions):
        payload = json.dumps(actions, ensure_ascii=False)
        raw = self.app.Run(f"{self.macro_module}.ExecuteActionsJson", payload)
        return json.loads(raw)


# ---- Action Dispatcher (for pywin32 backend) ----

def dispatch_action(engine, action):
    """Dispatch a single JSON action dict to the pywin32 engine."""
    act = action.get("action")
    sheet = action.get("sheet", 1)
    params = action.get("params", {})

    if act == "write_cell":
        engine.write_cell(sheet, action["cell"], action["value"], action.get("kind", "auto"))
        return {"ok": True, "action": act}

    elif act == "read_cell":
        val = engine.read_cell(sheet, action["cell"], action.get("property", "value2"))
        return {"ok": True, "action": act, "value": val}

    elif act == "write_range":
        engine.write_range(sheet, action["range"], action["values"])
        return {"ok": True, "action": act}

    elif act == "read_range":
        val = engine.read_range(sheet, action["range"])
        return {"ok": True, "action": act, "values": val}

    elif act == "clear_range":
        engine.clear_range(sheet, action["range"], params.get("mode", "contents"))
        return {"ok": True, "action": act}

    elif act == "merge_cells":
        engine.merge_cells(sheet, action["range"])
        return {"ok": True, "action": act}

    elif act == "unmerge_cells":
        engine.unmerge_cells(sheet, action["range"])
        return {"ok": True, "action": act}

    elif act == "set_format":
        engine.set_format(sheet, action["range"], **params)
        return {"ok": True, "action": act}

    elif act == "insert_rows":
        engine.insert_rows(sheet, action["row"], params.get("count", 1))
        return {"ok": True, "action": act}

    elif act == "delete_rows":
        engine.delete_rows(sheet, action["row"], params.get("count", 1))
        return {"ok": True, "action": act}

    elif act == "insert_cols":
        engine.insert_cols(sheet, action["col"], params.get("count", 1))
        return {"ok": True, "action": act}

    elif act == "delete_cols":
        engine.delete_cols(sheet, action["col"], params.get("count", 1))
        return {"ok": True, "action": act}

    elif act == "autofit_columns":
        engine.autofit_columns(sheet, action.get("range"))
        return {"ok": True, "action": act}

    elif act == "add_sheet":
        name = engine.add_sheet(params.get("name"), params.get("after"))
        return {"ok": True, "action": act, "name": name}

    elif act == "rename_sheet":
        engine.rename_sheet(action["old_name"], action["new_name"])
        return {"ok": True, "action": act}

    elif act == "delete_sheet":
        engine.delete_sheet(sheet)
        return {"ok": True, "action": act}

    elif act == "sort_range":
        engine.sort_range(sheet, action["range"], action["key_col"], params.get("order", "asc"))
        return {"ok": True, "action": act}

    elif act == "auto_filter":
        engine.auto_filter(sheet, action["range"], params.get("field"), params.get("criteria"))
        return {"ok": True, "action": act}

    elif act == "find_replace":
        engine.find_replace(sheet, action["range"], action["find"], action["replace"])
        return {"ok": True, "action": act}

    elif act == "calculate":
        engine.calculate()
        return {"ok": True, "action": act}

    elif act == "export_pdf":
        engine.export_pdf(action["path"])
        return {"ok": True, "action": act}

    elif act == "export_image":
        engine.export_image(sheet, action["path"], params.get("width", 1024), params.get("height", 768))
        return {"ok": True, "action": act}

    elif act == "copy_values":
        engine.copy_values(sheet, action["src_range"], action["dst_range"])
        return {"ok": True, "action": act}

    elif act == "paste_special":
        engine.paste_special(sheet, action["src_range"], action["dst_range"], params.get("paste_type", "values"))
        return {"ok": True, "action": act}

    elif act == "auto_fill":
        engine.auto_fill(sheet, action["src_range"], action["fill_range"])
        return {"ok": True, "action": act}

    elif act == "set_border":
        engine.set_border(sheet, action["range"], params.get("style", 1),
                          params.get("color", 0), params.get("weight", 2), params.get("edges", "all"))
        return {"ok": True, "action": act}

    elif act == "add_conditional_format":
        engine.add_conditional_format(sheet, action["range"], params.get("rule_type", "cell_value"),
                                       params.get("operator", 3), params.get("value"),
                                       params.get("format"))
        return {"ok": True, "action": act}

    elif act == "clear_conditional_format":
        engine.clear_conditional_format(sheet, action["range"])
        return {"ok": True, "action": act}

    elif act == "set_row_height":
        engine.set_row_height(sheet, action["row"], action["height"])
        return {"ok": True, "action": act}

    elif act == "set_col_width":
        engine.set_col_width(sheet, action["col"], action["width"])
        return {"ok": True, "action": act}

    elif act == "group_rows":
        engine.group_rows(sheet, action["start_row"], action["end_row"])
        return {"ok": True, "action": act}

    elif act == "ungroup_rows":
        engine.ungroup_rows(sheet, action["start_row"], action["end_row"])
        return {"ok": True, "action": act}

    elif act == "group_cols":
        engine.group_cols(sheet, action["start_col"], action["end_col"])
        return {"ok": True, "action": act}

    elif act == "ungroup_cols":
        engine.ungroup_cols(sheet, action["start_col"], action["end_col"])
        return {"ok": True, "action": act}

    elif act == "hide_rows":
        engine.hide_rows(sheet, action["start_row"], action["end_row"])
        return {"ok": True, "action": act}

    elif act == "unhide_rows":
        engine.unhide_rows(sheet, action["start_row"], action["end_row"])
        return {"ok": True, "action": act}

    elif act == "hide_cols":
        engine.hide_cols(sheet, action["start_col"], action["end_col"])
        return {"ok": True, "action": act}

    elif act == "unhide_cols":
        engine.unhide_cols(sheet, action["start_col"], action["end_col"])
        return {"ok": True, "action": act}

    elif act == "copy_sheet":
        name = engine.copy_sheet(sheet, params.get("before"), params.get("after"))
        return {"ok": True, "action": act, "name": name}

    elif act == "move_sheet":
        engine.move_sheet(sheet, params.get("before"), params.get("after"))
        return {"ok": True, "action": act}

    elif act == "protect_sheet":
        engine.protect_sheet(sheet, params.get("password"))
        return {"ok": True, "action": act}

    elif act == "unprotect_sheet":
        engine.unprotect_sheet(sheet, params.get("password"))
        return {"ok": True, "action": act}

    elif act == "freeze_panes":
        engine.freeze_panes(sheet, action["row"], action["col"])
        return {"ok": True, "action": act}

    elif act == "unfreeze_panes":
        engine.unfreeze_panes()
        return {"ok": True, "action": act}

    elif act == "add_validation":
        engine.add_validation(sheet, action["range"], params.get("type", "list"),
                              params.get("formula"), params.get("values"))
        return {"ok": True, "action": act}

    elif act == "clear_validation":
        engine.clear_validation(sheet, action["range"])
        return {"ok": True, "action": act}

    elif act == "add_named_range":
        engine.add_named_range(action["name"], action["refers_to"])
        return {"ok": True, "action": act}

    elif act == "delete_named_range":
        engine.delete_named_range(action["name"])
        return {"ok": True, "action": act}

    elif act == "add_chart":
        name = engine.add_chart(sheet, params.get("chart_type", "xlColumnClustered"),
                                action.get("data_range"), params.get("left", 100),
                                params.get("top", 100), params.get("width", 400), params.get("height", 300))
        return {"ok": True, "action": act, "name": name}

    elif act == "delete_chart":
        engine.delete_chart(sheet, action["chart_name"])
        return {"ok": True, "action": act}

    elif act == "modify_chart":
        engine.modify_chart(sheet, action["chart_name"], params.get("chart_type", "xlColumnClustered"))
        return {"ok": True, "action": act}

    elif act == "set_chart_title":
        engine.set_chart_title(sheet, action["chart_name"], action["title"])
        return {"ok": True, "action": act}

    elif act == "add_picture":
        name = engine.add_picture(sheet, action["path"], params.get("left", 0),
                                  params.get("top", 0), params.get("width", -1), params.get("height", -1))
        return {"ok": True, "action": act, "name": name}

    elif act == "delete_picture":
        engine.delete_picture(sheet, action["name"])
        return {"ok": True, "action": act}

    elif act == "add_comment":
        engine.add_comment(sheet, action["cell"], action["text"], params.get("author"))
        return {"ok": True, "action": act}

    elif act == "delete_comment":
        engine.delete_comment(sheet, action["cell"])
        return {"ok": True, "action": act}

    elif act == "add_hyperlink":
        engine.add_hyperlink(sheet, action["cell"], action["url"], params.get("display_text"))
        return {"ok": True, "action": act}

    elif act == "delete_hyperlink":
        engine.delete_hyperlink(sheet, action["cell"])
        return {"ok": True, "action": act}

    elif act == "set_page_setup":
        engine.set_page_setup(sheet, **params)
        return {"ok": True, "action": act}

    elif act == "run_macro":
        result = engine.run_macro(action["macro_name"], *action.get("args", []))
        return {"ok": True, "action": act, "result": result}

    elif act == "add_pivot_table":
        name = engine.add_pivot_table(sheet, action["source_range"], action["dest_sheet"],
                                       action["dest_cell"], params.get("table_name", "PivotTable1"),
                                       params.get("rows"), params.get("cols"), params.get("values"))
        return {"ok": True, "action": act, "name": name}

    elif act == "refresh_pivot":
        engine.refresh_pivot(sheet, action["table_name"])
        return {"ok": True, "action": act}

    else:
        return {"ok": False, "error": f"Unknown action: {act}"}


# ---- CLI ----

def main():
    parser = argparse.ArgumentParser(description="Excel COM Editor")
    parser.add_argument("file", nargs="?", help="Excel file path")
    parser.add_argument("--create", action="store_true", help="Create new workbook")
    parser.add_argument("--inspect", action="store_true", help="Print workbook structure")
    parser.add_argument("--inspect-sheet", metavar="SHEET", help="Print sheet data")
    parser.add_argument("--exec-actions", metavar="JSON", help="Execute JSON actions")
    parser.add_argument("--exec-script", metavar="SCRIPT", help="Execute Python script")
    parser.add_argument("--interactive-actions", action="store_true", help="Interactive JSON session")
    parser.add_argument("--output", "-o", metavar="PATH", help="Save output path")
    parser.add_argument("--backend", choices=("pywin32", "vba"), default="pywin32")
    parser.add_argument("--headed", action="store_true", help="Visible Excel window")
    parser.add_argument("--max-rows", type=int, default=50)
    parser.add_argument("--max-cols", type=int, default=26)

    args = parser.parse_args()

    if not args.file:
        parser.print_help()
        sys.exit(1)

    visible = args.headed
    if args.backend == "vba":
        engine = ExcelVBA(visible=visible)
    else:
        engine = ExcelCOM(visible=visible)

    try:
        if args.create:
            engine.create(args.file)
        else:
            engine.open(args.file)

        if args.inspect:
            result = engine.inspect()
            print(json.dumps(result, ensure_ascii=False, indent=2))

        if args.inspect_sheet:
            result = engine.inspect_sheet(args.inspect_sheet, args.max_rows, args.max_cols)
            print(json.dumps(result, ensure_ascii=False, indent=2, default=str))

        if args.exec_actions:
            actions_input = args.exec_actions
            if os.path.isfile(actions_input):
                with open(actions_input, "r", encoding="utf-8") as f:
                    actions = json.load(f)
            else:
                actions = json.loads(actions_input)

            if not isinstance(actions, list):
                actions = [actions]

            if args.backend == "vba":
                results = engine.execute_actions(actions)
                print(json.dumps(results, ensure_ascii=False, indent=2, default=str))
            else:
                results = []
                for act in actions:
                    r = dispatch_action(engine, act)
                    results.append(r)
                print(json.dumps(results, ensure_ascii=False, indent=2, default=str))

        if args.exec_script:
            script_path = args.exec_script
            with open(script_path, "r", encoding="utf-8") as f:
                code = f.read()
            exec(code, {"excel": engine, "filepath": engine.filepath, "json": json, "os": os, "time": time})

        if args.interactive_actions:
            print("Interactive JSON session. Type actions or 'quit'.")
            while True:
                try:
                    line = input("> ").strip()
                except (EOFError, KeyboardInterrupt):
                    break
                if not line:
                    continue
                if line.lower() in ("quit", "exit"):
                    break
                if line.lower() == "inspect":
                    print(json.dumps(engine.inspect(), ensure_ascii=False, indent=2))
                    continue
                if line.lower() == "save":
                    engine.save(args.output)
                    print("Saved.")
                    continue
                try:
                    actions = json.loads(line)
                    if not isinstance(actions, list):
                        actions = [actions]
                    if args.backend == "vba":
                        results = engine.execute_actions(actions)
                    else:
                        results = [dispatch_action(engine, a) for a in actions]
                    print(json.dumps(results, ensure_ascii=False, indent=2, default=str))
                except json.JSONDecodeError as e:
                    print(f"JSON parse error: {e}")
                except Exception as e:
                    print(f"Error: {e}")

        if args.output:
            engine.save(args.output)

    finally:
        engine.close()


if __name__ == "__main__":
    main()
