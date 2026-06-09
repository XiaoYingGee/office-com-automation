"""Excel MCP Server (PyAddin backend) — uses in-process Python add-in bridge.

Same 6 tools as excel_mcp.py, but execute_code runs INSIDE Excel.exe
via the registered ExcelEditor.PyAddIn COM add-in (zero IPC for per-cell ops).

Prerequisites:
    1. Register add-in: python python_addin/excel_pyaddin.py
    2. Start Excel manually (add-ins only load in interactive sessions)
    3. Run this server: python mcp_server/excel_mcp_addin.py

Usage:
    python excel_mcp_addin.py              # MCP stdio transport
"""

import json
import os
import sys

try:
    from mcp.server.fastmcp import FastMCP
except ImportError:
    print("pip install mcp", file=sys.stderr)
    sys.exit(1)

import pythoncom
import win32com.client

# ---- Excel add-in connection ----

_app = None
_bridge = None


def _connect():
    global _app, _bridge
    if _bridge is not None:
        return
    pythoncom.CoInitialize()
    _app = win32com.client.Dispatch("Excel.Application")
    _app.Visible = True
    _bridge = _app.COMAddIns.Item("ExcelEditor.PyAddIn").Object
    ping = _bridge.Ping()
    if ping != "pong":
        raise RuntimeError(f"PyAddIn did not respond (got: {ping})")


def _ensure_workbook():
    _connect()
    if _app.Workbooks.Count == 0:
        raise RuntimeError("No workbook open. Use excel_open first.")


# ---- MCP Server ----

mcp = FastMCP("excel-addin",
              instructions="""Excel automation via in-process Python add-in (zero IPC).
Use excel_execute_code to run Python code inside Excel.exe.
Use excel_inspect to see workbook structure.""")


@mcp.tool()
def excel_open(path: str, create: bool = False) -> str:
    """Open or create an Excel workbook.

    Args:
        path: File path (.xlsx).
        create: If True, create new. Otherwise open existing.
    """
    _connect()
    abs_path = os.path.abspath(path)
    try:
        if create:
            wb = _app.Workbooks.Add()
            wb.SaveAs(abs_path, 51)
        else:
            _app.Workbooks.Open(abs_path)
        info = json.loads(_bridge.InspectJson())
        return json.dumps({"ok": True, "path": abs_path,
                          "sheets": [s["name"] for s in info["sheets"]]})
    except Exception as e:
        return json.dumps({"ok": False, "error": str(e)})


@mcp.tool()
def excel_save(path: str = None) -> str:
    """Save the current workbook. Optionally save-as to a new path."""
    try:
        _ensure_workbook()
        if path:
            _app.ActiveWorkbook.SaveAs(os.path.abspath(path), 51)
        else:
            _app.ActiveWorkbook.Save()
        return json.dumps({"ok": True})
    except Exception as e:
        return json.dumps({"ok": False, "error": str(e)})


@mcp.tool()
def excel_inspect() -> str:
    """Return workbook structure: sheets with names, used ranges, row/col counts."""
    try:
        _ensure_workbook()
        return _bridge.InspectJson()
    except Exception as e:
        return json.dumps({"ok": False, "error": str(e)})


@mcp.tool()
def excel_inspect_sheet(sheet: str = "1", max_rows: int = 50, max_cols: int = 26) -> str:
    """Return sheet data as 2D array."""
    try:
        _ensure_workbook()
        params = json.dumps({"sheet": sheet, "max_rows": max_rows, "max_cols": max_cols})
        return _bridge.InspectSheetJson(params)
    except Exception as e:
        return json.dumps({"ok": False, "error": str(e)})


@mcp.tool()
def excel_execute_code(code: str) -> str:
    """Execute Python code INSIDE the Excel process (zero IPC).

    Available variables:
        app  — Excel.Application (COM, in-process)
        wb   — ActiveWorkbook
        ws   — ActiveSheet
        result — set this to return data

    The code runs via the in-process Python add-in bridge. All COM object model
    access is in-process (no cross-process marshalling per property).

    Example:
        ws.Range("A1").Value2 = "Hello"
        ws.Range("A1").Font.Bold = True
        result = "done"

    Example (bulk):
        for i in range(1, 101):
            ws.Cells(i, 1).Value2 = i
            ws.Cells(i, 2).Value2 = i * i
        result = "wrote 100 rows"

    Returns: {"ok": true, "result": ...} or {"ok": false, "error": ..., "traceback": ...}
    """
    try:
        _ensure_workbook()
        return _bridge.ExecuteCode(code)
    except Exception as e:
        return json.dumps({"ok": False, "error": str(e)})


@mcp.tool()
def excel_execute_actions(actions: str) -> str:
    """Execute structured JSON actions via in-process add-in bridge.

    Args:
        actions: JSON string — single action or array of actions.
    """
    try:
        _ensure_workbook()
        return _bridge.ExecuteActionsJson(actions)
    except Exception as e:
        return json.dumps({"ok": False, "error": str(e)})


if __name__ == "__main__":
    mcp.run()
