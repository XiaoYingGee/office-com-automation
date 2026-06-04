// dispatch.rs - execute an OpRequest against Excel via COM.

use serde_json::json;

use crate::com::{bstr, i4, r8, vbool, Variant};
use crate::excel::ExcelApp;
use super::{ErrorCategory, ExcelError, OpRequest, OpResponse};

// ---------------------------------------------------------------------------
// COM error helper
// ---------------------------------------------------------------------------
fn map_com_error(e: &windows::core::Error) -> ExcelError {
    ExcelError {
        category: ErrorCategory::ComError,
        code: e.code().0 as i64,
        message: e.message().to_string(),
        hint: None,
    }
}

fn excel_err(category: ErrorCategory, message: impl Into<String>) -> ExcelError {
    ExcelError { category, code: 0, message: message.into(), hint: None }
}

fn excel_err_hint(category: ErrorCategory, message: impl Into<String>, hint: impl Into<String>) -> ExcelError {
    ExcelError { category, code: 0, message: message.into(), hint: Some(hint.into()) }
}

// ---------------------------------------------------------------------------
// Format map
// ---------------------------------------------------------------------------
fn format_code(fmt: &str) -> Result<i32, ExcelError> {
    match fmt {
        "xlsx" => Ok(51),
        "xlsm" => Ok(52),
        "xls"  => Ok(56),
        "csv"  => Ok(6),
        other  => Err(excel_err(ErrorCategory::UnsupportedFormat, format!("unknown format: {other}"))),
    }
}

// ---------------------------------------------------------------------------
// Sheet resolution helpers
// ---------------------------------------------------------------------------
/// Resolve sheet with fallback: on name miss, fall back to first sheet.
/// Used for cell.write (workbook may be newly created on non-English Excel).
fn resolve_sheet_write(excel: &ExcelApp, wb: &crate::com::Dispatch, sheet: &Option<String>) -> Result<crate::com::Dispatch, ExcelError> {
    match sheet {
        Some(name) => {
            excel.get_sheet_by_name(wb, name).or_else(|_| {
                // Fall back to first sheet (handles non-English Excel default names)
                excel.get_sheet(wb, 1).map_err(|e| map_com_error(&e))
            })
        }
        None => excel.get_sheet(wb, 1).map_err(|e| map_com_error(&e)),
    }
}

/// Resolve sheet strictly: name miss → SheetNotFound.
/// Used for range.read.
fn resolve_sheet_read(excel: &ExcelApp, wb: &crate::com::Dispatch, sheet: &Option<String>) -> Result<crate::com::Dispatch, ExcelError> {
    match sheet {
        Some(name) => {
            excel.get_sheet_by_name(wb, name).map_err(|_| {
                excel_err(ErrorCategory::SheetNotFound, format!("sheet not found: {name}"))
            })
        }
        None => excel.get_sheet(wb, 1).map_err(|e| map_com_error(&e)),
    }
}

/// Normalize a path to Windows backslash format for COM calls.
fn win_path(path: &str) -> String {
    path.replace('/', "\\")
}

// ---------------------------------------------------------------------------
// cell.write
// ---------------------------------------------------------------------------
fn cell_write(req: OpRequest) -> Result<serde_json::Value, ExcelError> {
    let range_addr = req.target.as_ref()
        .and_then(|t| t.range.as_deref())
        .ok_or_else(|| excel_err(ErrorCategory::InvalidArg, "cell.write requires target.range"))?;

    let sheet_name = req.target.as_ref().and_then(|t| t.sheet.clone());

    let kind = req.params.get("kind")
        .and_then(|v| v.as_str())
        .ok_or_else(|| excel_err(ErrorCategory::InvalidArg, "params.kind is required"))?
        .to_string();

    let excel = ExcelApp::new().map_err(|e| map_com_error(&e))?;
    let native_path = win_path(&req.path);
    let wb = excel.open_or_create(&native_path).map_err(|e| map_com_error(&e))?;

    let ws = resolve_sheet_write(&excel, &wb, &sheet_name)?;
    let cell = excel.get_range(&ws, range_addr).map_err(|e| map_com_error(&e))?;

    // Write value by kind
    let value = &req.params["value"];
    match kind.as_str() {
        "string" => {
            let owned;
            let s = if let Some(s) = value.as_str() {
                s
            } else {
                owned = value.to_string();
                &owned
            };
            cell.put("Value2", bstr(s)).map_err(|e| map_com_error(&e))?;
        }
        "number" => {
            let n = value.as_f64()
                .ok_or_else(|| excel_err(ErrorCategory::InvalidArg, "params.value must be a number for kind=number"))?;
            cell.put("Value2", r8(n)).map_err(|e| map_com_error(&e))?;
        }
        "bool" => {
            let b = value.as_bool()
                .ok_or_else(|| excel_err(ErrorCategory::InvalidArg, "params.value must be a bool for kind=bool"))?;
            cell.put("Value2", vbool(b)).map_err(|e| map_com_error(&e))?;
        }
        "formula" => {
            let s = value.as_str()
                .ok_or_else(|| excel_err(ErrorCategory::InvalidArg, "params.value must be a string for kind=formula"))?;
            cell.put("Formula", bstr(s)).map_err(|e| map_com_error(&e))?;
        }
        other => {
            return Err(excel_err(ErrorCategory::InvalidArg, format!("unknown kind: {other}")));
        }
    }

    // Save
    if let Some(save_as) = &req.save_as {
        let fmt = format_code(&save_as.format)?;
        let save_path = win_path(&save_as.path);
        wb.call("SaveAs", &[bstr(&save_path), i4(fmt)]).map_err(|e| map_com_error(&e))?;
    } else {
        // Derive format from req.path extension; default to xlsx (51) if unknown.
        let ext = std::path::Path::new(&req.path)
            .extension()
            .and_then(|e| e.to_str())
            .unwrap_or("");
        let fmt = format_code(ext).unwrap_or(51);
        let save_path = win_path(&req.path);
        wb.call("SaveAs", &[bstr(&save_path), i4(fmt)]).map_err(|e| map_com_error(&e))?;
    }

    wb.call("Close", &[vbool(false)]).map_err(|e| map_com_error(&e))?;

    Ok(json!({"written": true}))
}

// ---------------------------------------------------------------------------
// range.read
// ---------------------------------------------------------------------------
fn range_read(req: OpRequest) -> Result<serde_json::Value, ExcelError> {
    let range_addr = req.target.as_ref()
        .and_then(|t| t.range.as_deref())
        .ok_or_else(|| excel_err(ErrorCategory::InvalidArg, "range.read requires target.range"))?;

    let sheet_name = req.target.as_ref().and_then(|t| t.sheet.clone());

    // File must exist
    if !std::path::Path::new(&req.path).exists() {
        return Err(excel_err(ErrorCategory::FileNotFound, format!("file not found: {}", req.path)));
    }

    let excel = ExcelApp::new().map_err(|e| map_com_error(&e))?;
    let native_path = win_path(&req.path);
    let wb = excel.open_workbook(&native_path).map_err(|e| map_com_error(&e))?;

    let ws = resolve_sheet_read(&excel, &wb, &sheet_name)?;
    let cell = excel.get_range(&ws, range_addr).map_err(|e| map_com_error(&e))?;

    let val = cell.get("Value2").map_err(|e| map_com_error(&e))?;

    let result = match val {
        Variant::Bstr(s)  => json!({"kind": "string",  "value": s}),
        Variant::R8(n)    => json!({"kind": "number",  "value": n}),
        Variant::I4(n)    => json!({"kind": "number",  "value": n as f64}),
        Variant::Bool(b)  => json!({"kind": "bool",    "value": b}),
        Variant::Empty | Variant::Missing => json!({"kind": "empty", "value": null}),
        _                 => json!({"kind": "unknown", "value": null}),
    };

    wb.call("Close", &[vbool(false)]).map_err(|e| map_com_error(&e))?;

    Ok(result)
}

// ---------------------------------------------------------------------------
// Public entry point
// ---------------------------------------------------------------------------
pub fn execute(req: OpRequest) -> OpResponse {
    match req.op.as_str() {
        "cell.write" => match cell_write(req) {
            Ok(v)  => OpResponse::ok(v),
            Err(e) => OpResponse::err(e),
        },
        "range.read" => match range_read(req) {
            Ok(v)  => OpResponse::ok(v),
            Err(e) => OpResponse::err(e),
        },
        other => OpResponse::err(excel_err_hint(
            ErrorCategory::Unknown,
            format!("unknown op: {other}"),
            "supported: cell.write, range.read",
        )),
    }
}
