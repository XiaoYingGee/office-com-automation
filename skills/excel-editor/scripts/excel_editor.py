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
