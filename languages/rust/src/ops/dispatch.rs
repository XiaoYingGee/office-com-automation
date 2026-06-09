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
// save_and_close — shared helper used by all mutating ops
// ---------------------------------------------------------------------------
/// Save (optionally as a new path/format) then close the workbook.
/// If `req.save_as` is set, uses that path/format; otherwise derives the
/// format from `req.path`'s extension, falling back to xlsx (51).
fn save_and_close(wb: &crate::com::Dispatch, req: &OpRequest) -> Result<(), ExcelError> {
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
    Ok(())
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

    save_and_close(&wb, &req)?;

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
                // Non-string (numeric cell has no formula) — route through variant_to_json
                other => variant_to_json(other),
            }
        }
        "text" => {
            let val = cell.get("Text").map_err(|e| map_com_error(&e))?;
            match val {
                Variant::Bstr(s) => json!({"kind": "string", "value": s}),
                Variant::Empty | Variant::Missing => json!({"kind": "empty", "value": null}),
                // Non-string variant — route through variant_to_json
                other => variant_to_json(other),
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

    // Validate that EVERY row has exactly `cols` elements (ragged-array guard)
    for (ri, row) in values.iter().enumerate() {
        let row_arr = row.as_array()
            .ok_or_else(|| excel_err(ErrorCategory::InvalidArg,
                format!("row {ri} in params.values is not an array")))?;
        if row_arr.len() as i32 != cols {
            return Err(excel_err(ErrorCategory::InvalidArg,
                format!("row {ri} has {} elements but row 0 has {cols}; all rows must have the same length",
                    row_arr.len())));
        }
    }

    // Build SafeArray2D
    let mut sa = SafeArray2D::new(rows, cols).map_err(|e| map_com_error(&e))?;
    for (ri, row) in values.iter().enumerate() {
        let row_arr = row.as_array().unwrap(); // already validated above
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

    save_and_close(&wb, &req)?;

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

    // File must exist (setup phase ensures it does)
    if !std::path::Path::new(&req.path).exists() {
        return Err(excel_err(ErrorCategory::FileNotFound, format!("file not found: {}", req.path)));
    }

    let excel = ExcelApp::new().map_err(|e| map_com_error(&e))?;
    let native_path = win_path(&req.path);
    let wb = excel.open_workbook(&native_path).map_err(|e| map_com_error(&e))?;

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

    save_and_close(&wb, &req)?;

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

    // File must exist (setup phase ensures it does)
    if !std::path::Path::new(&req.path).exists() {
        return Err(excel_err(ErrorCategory::FileNotFound, format!("file not found: {}", req.path)));
    }

    let excel = ExcelApp::new().map_err(|e| map_com_error(&e))?;
    let native_path = win_path(&req.path);
    let wb = excel.open_workbook(&native_path).map_err(|e| map_com_error(&e))?;

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

    save_and_close(&wb, &req)?;

    Ok(json!({"copied": true}))
}

// ---------------------------------------------------------------------------
// inspect
// ---------------------------------------------------------------------------
fn inspect(req: OpRequest) -> Result<serde_json::Value, ExcelError> {
    if !std::path::Path::new(&req.path).exists() {
        return Err(excel_err(ErrorCategory::FileNotFound, format!("file not found: {}", req.path)));
    }

    let excel = ExcelApp::new().map_err(|e| map_com_error(&e))?;
    let native_path = win_path(&req.path);
    let wb = excel.open_workbook(&native_path).map_err(|e| map_com_error(&e))?;

    let sheets_coll = wb.get_dispatch("Worksheets").map_err(|e| map_com_error(&e))?;
    let count = sheets_coll.get("Count").map_err(|e| map_com_error(&e))?.as_f64().unwrap_or(0.0) as i32;

    let mut sheets = Vec::new();
    for i in 1..=count {
        let ws = excel.get_sheet(&wb, i).map_err(|e| map_com_error(&e))?;
        let name = ws.get("Name").map_err(|e| map_com_error(&e))?;
        let name_str = match name { Variant::Bstr(s) => s, _ => String::new() };

        let used = ws.get_dispatch("UsedRange").map_err(|e| map_com_error(&e))?;
        let addr = used.get("Address").map_err(|e| map_com_error(&e))?;
        let addr_str = match addr { Variant::Bstr(s) => s, _ => String::new() };
        let rows = used.get_dispatch("Rows").map_err(|e| map_com_error(&e))?
            .get("Count").map_err(|e| map_com_error(&e))?.as_f64().unwrap_or(0.0) as i32;
        let cols = used.get_dispatch("Columns").map_err(|e| map_com_error(&e))?
            .get("Count").map_err(|e| map_com_error(&e))?.as_f64().unwrap_or(0.0) as i32;

        sheets.push(json!({
            "index": i,
            "name": name_str,
            "used_range": addr_str,
            "rows": rows,
            "cols": cols,
        }));
    }

    wb.call("Close", &[vbool(false)]).map_err(|e| map_com_error(&e))?;
    Ok(json!({"sheets": sheets}))
}

// ---------------------------------------------------------------------------
// set_format
// ---------------------------------------------------------------------------
fn set_format(req: OpRequest) -> Result<serde_json::Value, ExcelError> {
    let range_addr = req.target.as_ref()
        .and_then(|t| t.range.as_deref())
        .ok_or_else(|| excel_err(ErrorCategory::InvalidArg, "set_format requires target.range"))?;
    let sheet_name = req.target.as_ref().and_then(|t| t.sheet.clone());

    if !std::path::Path::new(&req.path).exists() {
        return Err(excel_err(ErrorCategory::FileNotFound, format!("file not found: {}", req.path)));
    }

    let excel = ExcelApp::new().map_err(|e| map_com_error(&e))?;
    let native_path = win_path(&req.path);
    let wb = excel.open_workbook(&native_path).map_err(|e| map_com_error(&e))?;
    let ws = resolve_sheet_write(&excel, &wb, &sheet_name)?;
    let rng = excel.get_range(&ws, range_addr).map_err(|e| map_com_error(&e))?;

    if let Some(v) = req.params.get("bold").and_then(|v| v.as_bool()) {
        let font = rng.get_dispatch("Font").map_err(|e| map_com_error(&e))?;
        font.put("Bold", Variant::Bool(v)).map_err(|e| map_com_error(&e))?;
    }
    if let Some(v) = req.params.get("italic").and_then(|v| v.as_bool()) {
        let font = rng.get_dispatch("Font").map_err(|e| map_com_error(&e))?;
        font.put("Italic", Variant::Bool(v)).map_err(|e| map_com_error(&e))?;
    }
    if let Some(v) = req.params.get("font_size").and_then(|v| v.as_f64()) {
        let font = rng.get_dispatch("Font").map_err(|e| map_com_error(&e))?;
        font.put("Size", Variant::R8(v)).map_err(|e| map_com_error(&e))?;
    }
    if let Some(v) = req.params.get("font_name").and_then(|v| v.as_str()) {
        let font = rng.get_dispatch("Font").map_err(|e| map_com_error(&e))?;
        font.put("Name", Variant::Bstr(v.to_string())).map_err(|e| map_com_error(&e))?;
    }
    if let Some(v) = req.params.get("font_color").and_then(|v| v.as_i64()) {
        let font = rng.get_dispatch("Font").map_err(|e| map_com_error(&e))?;
        font.put("Color", Variant::I4(v as i32)).map_err(|e| map_com_error(&e))?;
    }
    if let Some(v) = req.params.get("bg_color").and_then(|v| v.as_i64()) {
        let interior = rng.get_dispatch("Interior").map_err(|e| map_com_error(&e))?;
        interior.put("Color", Variant::I4(v as i32)).map_err(|e| map_com_error(&e))?;
    }
    if let Some(v) = req.params.get("number_format").and_then(|v| v.as_str()) {
        rng.put("NumberFormat", Variant::Bstr(v.to_string())).map_err(|e| map_com_error(&e))?;
    }

    save_and_close(&wb, &req)?;
    Ok(json!({"formatted": true}))
}

// ---------------------------------------------------------------------------
// row.insert
// ---------------------------------------------------------------------------
fn row_insert(req: OpRequest) -> Result<serde_json::Value, ExcelError> {
    let sheet_name = req.target.as_ref().and_then(|t| t.sheet.clone());
    let row = req.params.get("row").and_then(|v| v.as_i64())
        .ok_or_else(|| excel_err(ErrorCategory::InvalidArg, "row.insert requires params.row"))? as i32;
    let count = req.params.get("count").and_then(|v| v.as_i64()).unwrap_or(1) as i32;

    if !std::path::Path::new(&req.path).exists() {
        return Err(excel_err(ErrorCategory::FileNotFound, format!("file not found: {}", req.path)));
    }

    let excel = ExcelApp::new().map_err(|e| map_com_error(&e))?;
    let native_path = win_path(&req.path);
    let wb = excel.open_workbook(&native_path).map_err(|e| map_com_error(&e))?;
    let ws = resolve_sheet_write(&excel, &wb, &sheet_name)?;

    let rows_coll = ws.get_dispatch("Rows").map_err(|e| map_com_error(&e))?;
    for _ in 0..count {
        let row_obj = rows_coll.get_dispatch_with_args("Item", &[i4(row)]).map_err(|e| map_com_error(&e))?;
        row_obj.call("Insert", &[]).map_err(|e| map_com_error(&e))?;
    }

    save_and_close(&wb, &req)?;
    Ok(json!({"inserted": true, "row": row, "count": count}))
}

// ---------------------------------------------------------------------------
// row.delete
// ---------------------------------------------------------------------------
fn row_delete(req: OpRequest) -> Result<serde_json::Value, ExcelError> {
    let sheet_name = req.target.as_ref().and_then(|t| t.sheet.clone());
    let row = req.params.get("row").and_then(|v| v.as_i64())
        .ok_or_else(|| excel_err(ErrorCategory::InvalidArg, "row.delete requires params.row"))? as i32;
    let count = req.params.get("count").and_then(|v| v.as_i64()).unwrap_or(1) as i32;

    if !std::path::Path::new(&req.path).exists() {
        return Err(excel_err(ErrorCategory::FileNotFound, format!("file not found: {}", req.path)));
    }

    let excel = ExcelApp::new().map_err(|e| map_com_error(&e))?;
    let native_path = win_path(&req.path);
    let wb = excel.open_workbook(&native_path).map_err(|e| map_com_error(&e))?;
    let ws = resolve_sheet_write(&excel, &wb, &sheet_name)?;

    let addr = format!("{}:{}", row, row + count - 1);
    let rows_rng = ws.get_dispatch_with_args("Rows", &[bstr(&addr)]).map_err(|e| map_com_error(&e))?;
    rows_rng.call("Delete", &[]).map_err(|e| map_com_error(&e))?;

    save_and_close(&wb, &req)?;
    Ok(json!({"deleted": true, "row": row, "count": count}))
}

// ---------------------------------------------------------------------------
// sheet.add
// ---------------------------------------------------------------------------
fn sheet_add(req: OpRequest) -> Result<serde_json::Value, ExcelError> {
    let name = req.params.get("name").and_then(|v| v.as_str());

    let excel = ExcelApp::new().map_err(|e| map_com_error(&e))?;
    let native_path = win_path(&req.path);
    let wb = excel.open_or_create(&native_path).map_err(|e| map_com_error(&e))?;

    let new_ws = excel.add_sheet(&wb).map_err(|e| map_com_error(&e))?;
    if let Some(n) = name {
        new_ws.put("Name", Variant::Bstr(n.to_string())).map_err(|e| map_com_error(&e))?;
    }
    let actual_name = new_ws.get("Name").map_err(|e| map_com_error(&e))?;
    let name_str = match actual_name { Variant::Bstr(s) => s, _ => String::new() };

    save_and_close(&wb, &req)?;
    Ok(json!({"added": true, "name": name_str}))
}

// ---------------------------------------------------------------------------
// sheet.rename
// ---------------------------------------------------------------------------
fn sheet_rename(req: OpRequest) -> Result<serde_json::Value, ExcelError> {
    let sheet_name = req.target.as_ref().and_then(|t| t.sheet.clone())
        .ok_or_else(|| excel_err(ErrorCategory::InvalidArg, "sheet.rename requires target.sheet"))?;
    let new_name = req.params.get("name").and_then(|v| v.as_str())
        .ok_or_else(|| excel_err(ErrorCategory::InvalidArg, "sheet.rename requires params.name"))?;

    if !std::path::Path::new(&req.path).exists() {
        return Err(excel_err(ErrorCategory::FileNotFound, format!("file not found: {}", req.path)));
    }

    let excel = ExcelApp::new().map_err(|e| map_com_error(&e))?;
    let native_path = win_path(&req.path);
    let wb = excel.open_workbook(&native_path).map_err(|e| map_com_error(&e))?;

    let ws = excel.get_sheet_by_name(&wb, &sheet_name).map_err(|_| {
        excel_err(ErrorCategory::SheetNotFound, format!("sheet not found: {sheet_name}"))
    })?;
    ws.put("Name", Variant::Bstr(new_name.to_string())).map_err(|e| map_com_error(&e))?;

    save_and_close(&wb, &req)?;
    Ok(json!({"renamed": true, "old_name": sheet_name, "new_name": new_name}))
}

// ---------------------------------------------------------------------------
// sheet.delete
// ---------------------------------------------------------------------------
fn sheet_delete(req: OpRequest) -> Result<serde_json::Value, ExcelError> {
    let sheet_name = req.target.as_ref().and_then(|t| t.sheet.clone())
        .ok_or_else(|| excel_err(ErrorCategory::InvalidArg, "sheet.delete requires target.sheet"))?;

    if !std::path::Path::new(&req.path).exists() {
        return Err(excel_err(ErrorCategory::FileNotFound, format!("file not found: {}", req.path)));
    }

    let excel = ExcelApp::new().map_err(|e| map_com_error(&e))?;
    let native_path = win_path(&req.path);
    let wb = excel.open_workbook(&native_path).map_err(|e| map_com_error(&e))?;

    let ws = excel.get_sheet_by_name(&wb, &sheet_name).map_err(|_| {
        excel_err(ErrorCategory::SheetNotFound, format!("sheet not found: {sheet_name}"))
    })?;
    ws.call("Delete", &[]).map_err(|e| map_com_error(&e))?;

    save_and_close(&wb, &req)?;
    Ok(json!({"deleted": true, "name": sheet_name}))
}

// ---------------------------------------------------------------------------
// Public entry point
// ---------------------------------------------------------------------------
pub fn execute(req: OpRequest) -> OpResponse {
    let op = req.op.as_str();
    let result = match op {
        "cell.write"        => cell_write(req),
        "range.read"        => range_read(req),
        "range.write_bulk"  => range_write_bulk(req),
        "range.clear"       => range_clear(req),
        "range.copy_values" => range_copy_values(req),
        "inspect"           => inspect(req),
        "set_format"        => set_format(req),
        "row.insert"        => row_insert(req),
        "row.delete"        => row_delete(req),
        "sheet.add"         => sheet_add(req),
        "sheet.rename"      => sheet_rename(req),
        "sheet.delete"      => sheet_delete(req),
        other => return OpResponse::err(excel_err_hint(
            ErrorCategory::Unknown,
            format!("unknown op: {other}"),
            "supported: cell.write, range.read, range.write_bulk, range.clear, range.copy_values, inspect, set_format, row.insert, row.delete, sheet.add, sheet.rename, sheet.delete",
        )),
    };
    match result {
        Ok(v)  => OpResponse::ok(v),
        Err(e) => OpResponse::err(e),
    }
}
