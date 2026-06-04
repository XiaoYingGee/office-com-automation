// dispatch.rs - execute an OpRequest against Excel via COM.

use serde_json::json;

use crate::com::{bstr, i4, r8, vbool, SafeArray2D, Variant};
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

    // Optional params.property: "value2" (default) | "formula" | "text"
    let property = req.params.get("property")
        .and_then(|v| v.as_str())
        .unwrap_or("value2")
        .to_string();

    // File must exist
    if !std::path::Path::new(&req.path).exists() {
        return Err(excel_err(ErrorCategory::FileNotFound, format!("file not found: {}", req.path)));
    }

    let excel = ExcelApp::new().map_err(|e| map_com_error(&e))?;
    let native_path = win_path(&req.path);
    let wb = excel.open_workbook(&native_path).map_err(|e| map_com_error(&e))?;

    let ws = resolve_sheet_read(&excel, &wb, &sheet_name)?;
    let cell = excel.get_range(&ws, range_addr).map_err(|e| map_com_error(&e))?;

    let result = match property.as_str() {
        "value2" => {
            let val = cell.get("Value2").map_err(|e| map_com_error(&e))?;
            variant_to_json(val)
        }
        "formula" => {
            let val = cell.get("Formula").map_err(|e| map_com_error(&e))?;
            match val {
                Variant::Bstr(s) => json!({"kind": "string", "value": s}),
                Variant::Empty | Variant::Missing => json!({"kind": "empty", "value": null}),
                _ => json!({"kind": "string", "value": null}),
            }
        }
        "text" => {
            let val = cell.get("Text").map_err(|e| map_com_error(&e))?;
            match val {
                Variant::Bstr(s) => json!({"kind": "string", "value": s}),
                Variant::Empty | Variant::Missing => json!({"kind": "empty", "value": null}),
                _ => json!({"kind": "string", "value": null}),
            }
        }
        other => {
            wb.call("Close", &[vbool(false)]).map_err(|e| map_com_error(&e))?;
            return Err(excel_err(ErrorCategory::InvalidArg,
                format!("unknown params.property: {other}; expected value2|formula|text")));
        }
    };

    wb.call("Close", &[vbool(false)]).map_err(|e| map_com_error(&e))?;

    Ok(result)
}

// Map a Variant to a JSON {kind, value} object (for Value2 reads).
fn variant_to_json(val: Variant) -> serde_json::Value {
    match val {
        Variant::Bstr(s)  => json!({"kind": "string",  "value": s}),
        Variant::R8(n)    => json!({"kind": "number",  "value": n}),
        Variant::I4(n)    => json!({"kind": "number",  "value": n as f64}),
        Variant::Bool(b)  => json!({"kind": "bool",    "value": b}),
        Variant::Empty | Variant::Missing => json!({"kind": "empty", "value": null}),
        _                 => json!({"kind": "unknown", "value": null}),
    }
}

// ---------------------------------------------------------------------------
// range.write_bulk
// ---------------------------------------------------------------------------
fn range_write_bulk(req: OpRequest) -> Result<serde_json::Value, ExcelError> {
    let range_addr = req.target.as_ref()
        .and_then(|t| t.range.as_deref())
        .ok_or_else(|| excel_err(ErrorCategory::InvalidArg, "range.write_bulk requires target.range"))?;

    // Validate that the range contains a colon (multi-cell range required)
    if !range_addr.contains(':') {
        return Err(excel_err(ErrorCategory::InvalidArg,
            "range.write_bulk requires a multi-cell target.range (e.g. \"A1:B2\")"));
    }

    let sheet_name = req.target.as_ref().and_then(|t| t.sheet.clone());

    // params.values must be a 2D JSON array
    let values = req.params.get("values")
        .and_then(|v| v.as_array())
        .ok_or_else(|| excel_err(ErrorCategory::InvalidArg, "range.write_bulk requires params.values (2D array)"))?;

    let rows = values.len() as i32;
    if rows == 0 {
        return Err(excel_err(ErrorCategory::InvalidArg, "params.values must not be empty"));
    }
    let cols = values[0].as_array()
        .map(|r| r.len())
        .unwrap_or(0) as i32;
    if cols == 0 {
        return Err(excel_err(ErrorCategory::InvalidArg, "params.values rows must not be empty"));
    }

    // Build SafeArray2D
    let mut sa = SafeArray2D::new(rows, cols).map_err(|e| map_com_error(&e))?;
    for (ri, row) in values.iter().enumerate() {
        let row_arr = row.as_array()
            .ok_or_else(|| excel_err(ErrorCategory::InvalidArg, "each row in params.values must be an array"))?;
        for (ci, cell_val) in row_arr.iter().enumerate() {
            let variant = json_scalar_to_variant(cell_val)?;
            sa.put((ri + 1) as i32, (ci + 1) as i32, variant)
                .map_err(|e| map_com_error(&e))?;
        }
    }
    let sa_variant = sa.into_variant();

    let excel = ExcelApp::new().map_err(|e| map_com_error(&e))?;
    let native_path = win_path(&req.path);
    let wb = excel.open_or_create(&native_path).map_err(|e| map_com_error(&e))?;

    let ws = resolve_sheet_write(&excel, &wb, &sheet_name)?;
    let rng = excel.get_range(&ws, range_addr).map_err(|e| map_com_error(&e))?;
    rng.put_raw("Value2", sa_variant).map_err(|e| map_com_error(&e))?;

    // Save
    if let Some(save_as) = &req.save_as {
        let fmt = format_code(&save_as.format)?;
        let save_path = win_path(&save_as.path);
        wb.call("SaveAs", &[bstr(&save_path), i4(fmt)]).map_err(|e| map_com_error(&e))?;
    } else {
        let ext = std::path::Path::new(&req.path)
            .extension()
            .and_then(|e| e.to_str())
            .unwrap_or("");
        let fmt = format_code(ext).unwrap_or(51);
        let save_path = win_path(&req.path);
        wb.call("SaveAs", &[bstr(&save_path), i4(fmt)]).map_err(|e| map_com_error(&e))?;
    }

    wb.call("Close", &[vbool(false)]).map_err(|e| map_com_error(&e))?;

    Ok(json!({"written": true, "rows": rows, "cols": cols}))
}

fn json_scalar_to_variant(v: &serde_json::Value) -> Result<Variant, ExcelError> {
    match v {
        serde_json::Value::String(s) => Ok(Variant::Bstr(s.clone())),
        serde_json::Value::Number(n) => {
            let f = n.as_f64().ok_or_else(|| excel_err(ErrorCategory::InvalidArg, "numeric value out of f64 range"))?;
            Ok(Variant::R8(f))
        }
        serde_json::Value::Bool(b) => Ok(Variant::Bool(*b)),
        serde_json::Value::Null => Ok(Variant::Empty),
        _ => Err(excel_err(ErrorCategory::InvalidArg, "params.values cells must be string, number, bool, or null")),
    }
}

// ---------------------------------------------------------------------------
// range.clear
// ---------------------------------------------------------------------------
fn range_clear(req: OpRequest) -> Result<serde_json::Value, ExcelError> {
    let range_addr = req.target.as_ref()
        .and_then(|t| t.range.as_deref())
        .ok_or_else(|| excel_err(ErrorCategory::InvalidArg, "range.clear requires target.range"))?;

    let sheet_name = req.target.as_ref().and_then(|t| t.sheet.clone());

    let mode = req.params.get("mode")
        .and_then(|v| v.as_str())
        .unwrap_or("contents")
        .to_string();

    // File must exist for clear to make sense; open_or_create to handle fresh files too
    let excel = ExcelApp::new().map_err(|e| map_com_error(&e))?;
    let native_path = win_path(&req.path);
    let wb = excel.open_or_create(&native_path).map_err(|e| map_com_error(&e))?;

    let ws = resolve_sheet_write(&excel, &wb, &sheet_name)?;
    let rng = excel.get_range(&ws, range_addr).map_err(|e| map_com_error(&e))?;

    match mode.as_str() {
        "contents" => {
            rng.call("ClearContents", &[]).map_err(|e| map_com_error(&e))?;
        }
        "all" => {
            rng.call("Clear", &[]).map_err(|e| map_com_error(&e))?;
        }
        other => {
            wb.call("Close", &[vbool(false)]).map_err(|e| map_com_error(&e))?;
            return Err(excel_err(ErrorCategory::InvalidArg,
                format!("unknown mode: {other}; expected contents|all")));
        }
    }

    // Save
    if let Some(save_as) = &req.save_as {
        let fmt = format_code(&save_as.format)?;
        let save_path = win_path(&save_as.path);
        wb.call("SaveAs", &[bstr(&save_path), i4(fmt)]).map_err(|e| map_com_error(&e))?;
    } else {
        let ext = std::path::Path::new(&req.path)
            .extension()
            .and_then(|e| e.to_str())
            .unwrap_or("");
        let fmt = format_code(ext).unwrap_or(51);
        let save_path = win_path(&req.path);
        wb.call("SaveAs", &[bstr(&save_path), i4(fmt)]).map_err(|e| map_com_error(&e))?;
    }

    wb.call("Close", &[vbool(false)]).map_err(|e| map_com_error(&e))?;

    Ok(json!({"cleared": true}))
}

// ---------------------------------------------------------------------------
// range.copy_values
// ---------------------------------------------------------------------------
fn range_copy_values(req: OpRequest) -> Result<serde_json::Value, ExcelError> {
    let src_addr = req.target.as_ref()
        .and_then(|t| t.range.as_deref())
        .ok_or_else(|| excel_err(ErrorCategory::InvalidArg, "range.copy_values requires target.range (source)"))?;

    let dest_addr = req.params.get("dest")
        .and_then(|v| v.as_str())
        .ok_or_else(|| excel_err(ErrorCategory::InvalidArg, "range.copy_values requires params.dest (destination address)"))?
        .to_string();

    let sheet_name = req.target.as_ref().and_then(|t| t.sheet.clone());

    let excel = ExcelApp::new().map_err(|e| map_com_error(&e))?;
    let native_path = win_path(&req.path);
    let wb = excel.open_or_create(&native_path).map_err(|e| map_com_error(&e))?;

    let ws = resolve_sheet_write(&excel, &wb, &sheet_name)?;

    // Read source Value2
    let src_rng = excel.get_range(&ws, src_addr).map_err(|e| map_com_error(&e))?;
    let src_val = src_rng.get("Value2").map_err(|e| map_com_error(&e))?;

    // Write to destination
    let dest_rng = excel.get_range(&ws, &dest_addr).map_err(|e| map_com_error(&e))?;
    match &src_val {
        Variant::Bstr(s) => dest_rng.put("Value2", Variant::Bstr(s.clone())).map_err(|e| map_com_error(&e))?,
        Variant::R8(n)   => dest_rng.put("Value2", Variant::R8(*n)).map_err(|e| map_com_error(&e))?,
        Variant::I4(n)   => dest_rng.put("Value2", Variant::R8(*n as f64)).map_err(|e| map_com_error(&e))?,
        Variant::Bool(b) => dest_rng.put("Value2", Variant::Bool(*b)).map_err(|e| map_com_error(&e))?,
        Variant::Empty | Variant::Missing => {
            // Source is empty; clearing the destination is equivalent — no-op is fine
        }
        _ => {
            wb.call("Close", &[vbool(false)]).map_err(|e| map_com_error(&e))?;
            return Err(excel_err(ErrorCategory::InvalidArg, "source cell has unsupported value type for copy_values"));
        }
    }

    // Save
    if let Some(save_as) = &req.save_as {
        let fmt = format_code(&save_as.format)?;
        let save_path = win_path(&save_as.path);
        wb.call("SaveAs", &[bstr(&save_path), i4(fmt)]).map_err(|e| map_com_error(&e))?;
    } else {
        let ext = std::path::Path::new(&req.path)
            .extension()
            .and_then(|e| e.to_str())
            .unwrap_or("");
        let fmt = format_code(ext).unwrap_or(51);
        let save_path = win_path(&req.path);
        wb.call("SaveAs", &[bstr(&save_path), i4(fmt)]).map_err(|e| map_com_error(&e))?;
    }

    wb.call("Close", &[vbool(false)]).map_err(|e| map_com_error(&e))?;

    Ok(json!({"copied": true}))
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
        "range.write_bulk" => match range_write_bulk(req) {
            Ok(v)  => OpResponse::ok(v),
            Err(e) => OpResponse::err(e),
        },
        "range.clear" => match range_clear(req) {
            Ok(v)  => OpResponse::ok(v),
            Err(e) => OpResponse::err(e),
        },
        "range.copy_values" => match range_copy_values(req) {
            Ok(v)  => OpResponse::ok(v),
            Err(e) => OpResponse::err(e),
        },
        other => OpResponse::err(excel_err_hint(
            ErrorCategory::Unknown,
            format!("unknown op: {other}"),
            "supported: cell.write, range.read, range.write_bulk, range.clear, range.copy_values",
        )),
    }
}
