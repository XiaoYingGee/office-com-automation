"""Excel MCP Server — exposes Excel automation to AI agents via MCP protocol.

Architecture:
    AI Agent ──MCP (stdio JSON-RPC)──→ this server ──COM──→ Excel.exe

    For simple ops: use excel_execute_actions (JSON dispatch)
    For complex/bulk ops: use excel_execute_code (AI writes Python, runs in-process)

Tools:
    1. excel_inspect        — workbook structure (sheets, ranges)
    2. excel_inspect_sheet  — sheet data as 2D array
    3. excel_execute_code   — arbitrary Python executed in Excel process (zero IPC)
    4. excel_execute_actions— batch JSON actions (structured, validated)
    5. excel_open           — open/create a workbook
    6. excel_save           — save current workbook

Usage:
    python excel_mcp.py                    # start MCP server (stdio)
    python excel_mcp.py --transport sse    # SSE transport (optional)

Requirements:
    pip install pywin32 mcp
"""

import json
import os
import sys
import traceback

import pythoncom
import win32com.client

try:
    from mcp.server.fastmcp import FastMCP
except ImportError:
    print("pip install mcp  (requires mcp package)", file=sys.stderr)
    sys.exit(1)

# Add the excel_editor scripts to path for dispatch_action reuse
SCRIPT_DIR = os.path.normpath(os.path.join(os.path.dirname(__file__),
                                            "..", "skills", "excel-editor", "scripts"))
if SCRIPT_DIR not in sys.path:
    sys.path.insert(0, SCRIPT_DIR)

from excel_editor import ExcelCOM, dispatch_action

# ---- Excel connection ----

_excel: ExcelCOM = None


def _get_excel() -> ExcelCOM:
    global _excel
    if _excel is None or _excel.app is None:
        pythoncom.CoInitialize()
        _excel = ExcelCOM(visible=True)
    return _excel


def _ensure_workbook():
    excel = _get_excel()
    if excel.wb is None:
        try:
            excel.wb = excel.app.ActiveWorkbook
        except Exception:
            pass
    if excel.wb is None:
        raise RuntimeError("No workbook open. Use excel_open first.")
    return excel


# ---- MCP Server ----

mcp = FastMCP("excel",
              instructions="""Excel automation server. Use excel_inspect to see workbook structure,
excel_execute_code to run Python code in Excel process (most flexible),
or excel_execute_actions for structured JSON operations.""")


@mcp.tool()
def excel_open(path: str, create: bool = False) -> str:
    """Open or create an Excel workbook.

    Args:
        path: File path (.xlsx). Relative paths resolved from CWD.
        create: If True, create a new workbook at path. Otherwise open existing.

    Returns: JSON {"ok": true, "path": "...", "sheets": [...]}
    """
    excel = _get_excel()
    try:
        if create:
            excel.create(path)
        else:
            excel.open(path)
        info = excel.inspect()
        return json.dumps({"ok": True, "path": excel.filepath,
                          "sheets": [s["name"] for s in info["sheets"]]},
                         ensure_ascii=False)
    except Exception as e:
        return json.dumps({"ok": False, "error": str(e)})


@mcp.tool()
def excel_save(path: str = None) -> str:
    """Save the current workbook. Optionally save-as to a new path.

    Args:
        path: If provided, SaveAs to this path. Otherwise save in-place.

    Returns: JSON {"ok": true}
    """
    try:
        excel = _ensure_workbook()
        excel.save(path)
        return json.dumps({"ok": True})
    except Exception as e:
        return json.dumps({"ok": False, "error": str(e)})


@mcp.tool()
def excel_inspect() -> str:
    """Return workbook structure: list of sheets with names, used ranges, row/col counts.

    Returns: JSON {"sheets": [{"index": 1, "name": "Sheet1", "used_range": "$A$1:$E$100", "rows": 100, "cols": 5}, ...]}
    """
    try:
        excel = _ensure_workbook()
        return json.dumps(excel.inspect(), ensure_ascii=False)
    except Exception as e:
        return json.dumps({"ok": False, "error": str(e)})


@mcp.tool()
def excel_inspect_sheet(sheet: str = "1", max_rows: int = 50, max_cols: int = 26) -> str:
    """Return sheet data as a 2D array.

    Args:
        sheet: Sheet name or 1-based index (as string).
        max_rows: Maximum rows to return (default 50).
        max_cols: Maximum columns to return (default 26).

    Returns: JSON {"name": "Sheet1", "used_range": "...", "data": [[...], ...]}
    """
    try:
        excel = _ensure_workbook()
        try:
            sheet_ref = int(sheet)
        except ValueError:
            sheet_ref = sheet
        result = excel.inspect_sheet(sheet_ref, max_rows, max_cols)
        return json.dumps(result, ensure_ascii=False, default=str)
    except Exception as e:
        return json.dumps({"ok": False, "error": str(e)})


@mcp.tool()
def excel_execute_code(code: str) -> str:
    """Execute Python code inside the Excel process context.

    The code runs with these variables available:
        app  — Excel.Application COM object
        wb   — ActiveWorkbook
        ws   — ActiveSheet
        result — set this variable to return data to the caller

    The code can access the full Excel COM object model. All operations run
    in-process with zero IPC overhead. Set `result` to return data.

    Example:
        # Write values and format
        ws.Range("A1").Value2 = "Hello"
        ws.Range("A1").Font.Bold = True
        ws.Range("A1").Font.Color = 0x0000FF  # Red (BGR)
        result = "done"

    Example:
        # Read and process data
        data = ws.Range("A1:C10").Value2
        total = sum(row[2] for row in data if row[2])
        result = {"total": total, "rows": len(data)}

    Example:
        # Bulk operations (all in one call, zero IPC per cell)
        for i in range(1, 101):
            ws.Cells(i, 1).Value2 = i
            ws.Cells(i, 2).Value2 = i * i
        ws.Range("A1:B100").Columns.AutoFit()
        result = "wrote 100 rows"

    Returns: JSON {"ok": true, "result": ...} or {"ok": false, "error": "...", "traceback": "..."}
    """
    try:
        excel = _ensure_workbook()
        local_vars = {
            "app": excel.app,
            "wb": excel.wb,
            "ws": excel.wb.ActiveSheet,
            "result": None,
            "json": json,
            "os": os,
        }
        exec(code, {"__builtins__": __builtins__}, local_vars)
        return json.dumps({"ok": True, "result": local_vars.get("result")},
                         ensure_ascii=False, default=str)
    except Exception as e:
        return json.dumps({"ok": False, "error": str(e),
                          "traceback": traceback.format_exc()},
                         ensure_ascii=False)


@mcp.tool()
def excel_execute_actions(actions: str) -> str:
    """Execute one or more structured JSON actions.

    Args:
        actions: JSON string — a single action object or array of action objects.

    Each action has the form:
        {"action": "write_cell", "sheet": "Sheet1", "cell": "A1", "value": 42, "kind": "auto"}

    Available actions (62 total):
        Cell I/O: write_cell, read_cell, write_range, read_range
        Area: clear_range, merge_cells, unmerge_cells, copy_values, paste_special, auto_fill
        Format: set_format, set_border, add_conditional_format, clear_conditional_format
        Rows/Cols: insert_rows, delete_rows, insert_cols, delete_cols, autofit_columns,
                   set_row_height, set_col_width, group_rows, ungroup_rows, group_cols,
                   ungroup_cols, hide_rows, unhide_rows, hide_cols, unhide_cols
        Sheets: add_sheet, rename_sheet, delete_sheet, copy_sheet, move_sheet,
                protect_sheet, unprotect_sheet, freeze_panes, unfreeze_panes
        Data: sort_range, auto_filter, find_replace, calculate,
              add_validation, clear_validation, add_named_range, delete_named_range
        Charts: add_chart, delete_chart, modify_chart, set_chart_title
        Pictures: add_picture, delete_picture
        Comments: add_comment, delete_comment
        Hyperlinks: add_hyperlink, delete_hyperlink
        Export: export_pdf, export_image, set_page_setup
        Macros: run_macro
        Pivot: add_pivot_table, refresh_pivot

    Returns: JSON array of results [{ok, action, ...}, ...]
    """
    try:
        excel = _ensure_workbook()
        action_list = json.loads(actions)
        if isinstance(action_list, dict):
            action_list = [action_list]
        results = []
        for action in action_list:
            try:
                r = dispatch_action(excel, action)
            except Exception as e:
                r = {"ok": False, "error": str(e), "action": action.get("action", "")}
            results.append(r)
        return json.dumps(results, ensure_ascii=False, default=str)
    except json.JSONDecodeError as e:
        return json.dumps({"ok": False, "error": f"JSON parse error: {e}"})
    except Exception as e:
        return json.dumps({"ok": False, "error": str(e)})


# ---- Entry point ----

if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description="Excel MCP Server")
    parser.add_argument("--transport", choices=("stdio", "sse"), default="stdio")
    args = parser.parse_args()

    if args.transport == "sse":
        mcp.run(transport="sse")
    else:
        mcp.run()
