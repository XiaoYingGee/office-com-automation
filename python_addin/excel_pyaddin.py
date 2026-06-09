"""In-process Excel COM add-in implemented in pure Python (pywin32).

Runs the editing logic INSIDE EXCEL.EXE via a COM add-in so the Excel object
model is touched with ZERO cross-process marshalling — the same reason VBA is
far faster than out-of-process pywin32.

External drivers reach it via:
    app.COMAddIns.Item("ExcelEditor.PyAddIn").Object  -> the bridge
Then ONE coarse cross-process call per operation:
    bridge.Ping()
    bridge.InspectJson()
    bridge.ExecuteActionJson(json_string)

All per-cell/range work happens in-process by reusing the proven ExcelCOM +
dispatch_action logic, bound to the in-process Application.

Register / unregister (per-user, no admin):
    python excel_pyaddin.py                # register COM server + Excel add-in
    python excel_pyaddin.py --unregister   # remove both
"""

import json
import os
import sys

import pythoncom
import win32com.server.util
from win32com import universal

universal.RegisterInterfaces(
    "{AC0714F2-3D04-11D1-AE7D-00A0C90F26F4}", 0, 1, 0, ["_IDTExtensibility2"]
)

PYADDIN_CLSID = "{7A1B3C4D-5E6F-4A8B-9C0D-1E2F3A4B5C6D}"
PYADDIN_PROGID = "ExcelEditor.PyAddIn"

_LOG_PATH = os.path.join(
    os.environ.get("TEMP", os.path.expanduser("~")), "exceleditor_pyaddin.log"
)


def _log(msg):
    try:
        import datetime
        with open(_LOG_PATH, "a", encoding="utf-8") as fh:
            fh.write(datetime.datetime.now().strftime("%H:%M:%S.%f ") + str(msg) + "\n")
    except Exception:
        pass


def _import_engine():
    here = os.path.dirname(os.path.abspath(__file__))
    scripts_dir = os.path.normpath(os.path.join(here, "..", "skills", "excel-editor", "scripts"))
    if scripts_dir not in sys.path:
        sys.path.insert(0, scripts_dir)
    if here not in sys.path:
        sys.path.insert(0, here)
    from excel_editor import ExcelCOM, dispatch_action
    return ExcelCOM, dispatch_action


class ExcelEditorPyAddin:
    """In-process add-in + automation bridge."""

    _reg_clsid_ = PYADDIN_CLSID
    _reg_progid_ = PYADDIN_PROGID
    _reg_desc_ = "ExcelEditor Python in-process add-in"
    _reg_class_spec_ = "excel_pyaddin.ExcelEditorPyAddin"
    _reg_policy_spec_ = "win32com.server.policy.EventHandlerPolicy"

    _com_interfaces_ = ["_IDTExtensibility2"]
    _public_methods_ = [
        "Ping",
        "InspectJson",
        "InspectSheetJson",
        "ExecuteActionJson",
        "ExecuteActionsJson",
        "ExecuteCode",
        "RequestComAddInAutomationService",
    ]

    def __init__(self):
        self._app = None
        self._engine = None
        self._dispatch = None

    # ---- IDTExtensibility2 ----

    def OnConnection(self, application, connectMode, addin, custom):
        _log("OnConnection connectMode=%s" % connectMode)
        self._app = application
        try:
            ExcelCOM, dispatch_action = _import_engine()
            self._dispatch = dispatch_action
            engine = ExcelCOM.__new__(ExcelCOM)
            engine.app = application
            engine.wb = None
            engine.filepath = None
            self._engine = engine
            _log("engine bound OK")
        except Exception as exc:
            _log("engine bind FAILED: %s" % exc)
        try:
            addin.Object = win32com.server.util.wrap(self)
            _log("addin.Object set OK")
        except Exception as exc:
            _log("addin.Object set FAILED: %s" % exc)

    def OnDisconnection(self, mode, custom):
        _log("OnDisconnection mode=%s" % mode)
        self._engine = None
        self._app = None

    def OnAddInsUpdate(self, custom):
        pass

    def OnStartupComplete(self, custom):
        pass

    def OnBeginShutdown(self, custom):
        pass

    def RequestComAddInAutomationService(self):
        return win32com.server.util.wrap(self)

    # ---- Automation bridge ----

    def _bind_active(self):
        if self._engine is None:
            raise RuntimeError("engine not initialized")
        wb = None
        try:
            wb = self._app.ActiveWorkbook
        except Exception:
            wb = None
        if wb is None:
            try:
                wb = self._app.Workbooks(1)
            except Exception:
                pass
        self._engine.wb = wb
        try:
            self._engine.filepath = wb.FullName if wb else None
        except Exception:
            self._engine.filepath = None
        return self._engine

    def Ping(self):
        return "pong"

    def InspectJson(self):
        engine = self._bind_active()
        return json.dumps(engine.inspect(), ensure_ascii=False)

    def ExecuteActionJson(self, actionJson):
        engine = self._bind_active()
        action = json.loads(actionJson)
        result = self._dispatch(engine, action)
        return json.dumps(result, ensure_ascii=False, default=str)

    def ExecuteActionsJson(self, actionsJson):
        engine = self._bind_active()
        actions = json.loads(actionsJson)
        results = []
        for action in actions:
            try:
                r = self._dispatch(engine, action)
            except Exception as e:
                r = {"ok": False, "error": str(e), "action": action.get("action", "")}
            results.append(r)
        return json.dumps(results, ensure_ascii=False, default=str)

    def InspectSheetJson(self, paramsJson):
        engine = self._bind_active()
        params = json.loads(paramsJson) if paramsJson else {}
        sheet = params.get("sheet", 1)
        max_rows = params.get("max_rows", 50)
        max_cols = params.get("max_cols", 26)
        result = engine.inspect_sheet(sheet, max_rows, max_cols)
        return json.dumps(result, ensure_ascii=False, default=str)

    def ExecuteCode(self, code):
        """Execute arbitrary Python code in-process.

        Available variables:
            app  - Excel.Application (COM object, in-process)
            wb   - ActiveWorkbook
            ws   - ActiveSheet
            result - set this to return data to the caller

        Returns JSON: {"ok": true, "result": ...} or {"ok": false, "error": ..., "traceback": ...}
        """
        self._bind_active()
        local_vars = {
            "app": self._app,
            "wb": self._engine.wb,
            "ws": self._engine.wb.ActiveSheet if self._engine.wb else None,
            "result": None,
            "json": json,
            "os": os,
        }
        try:
            exec(code, {"__builtins__": __builtins__}, local_vars)
            return json.dumps({"ok": True, "result": local_vars.get("result")},
                              ensure_ascii=False, default=str)
        except Exception as e:
            import traceback
            return json.dumps({"ok": False, "error": str(e),
                              "traceback": traceback.format_exc()},
                             ensure_ascii=False)


# ---- register / unregister ----

_ADDINS_KEY = r"Software\Microsoft\Office\Excel\Addins"


def _register_addin(klass):
    import winreg
    with winreg.CreateKey(winreg.HKEY_CURRENT_USER, _ADDINS_KEY + "\\" + klass._reg_progid_) as sub:
        winreg.SetValueEx(sub, "CommandLineSafe", 0, winreg.REG_DWORD, 0)
        winreg.SetValueEx(sub, "LoadBehavior", 0, winreg.REG_DWORD, 3)
        winreg.SetValueEx(sub, "Description", 0, winreg.REG_SZ, klass._reg_desc_)
        winreg.SetValueEx(sub, "FriendlyName", 0, winreg.REG_SZ, "ExcelEditor Python Add-In")
    print("Registered '%s' (CLSID %s) per-user. LoadBehavior=3." % (klass._reg_progid_, klass._reg_clsid_))


def _unregister_addin(klass):
    import winreg
    try:
        winreg.DeleteKey(winreg.HKEY_CURRENT_USER, _ADDINS_KEY + "\\" + klass._reg_progid_)
        print("Removed Excel add-in key '%s'." % klass._reg_progid_)
    except OSError:
        pass


if __name__ == "__main__":
    import win32com.server.register
    win32com.server.register.UseCommandLine(ExcelEditorPyAddin)
    if "--unregister" in sys.argv or "--clean" in sys.argv:
        _unregister_addin(ExcelEditorPyAddin)
    else:
        _register_addin(ExcelEditorPyAddin)
