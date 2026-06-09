// ops.cpp - Backend-protocol op dispatch for the C++ COM backend.
//
// Reads one OpRequest (parsed JSON), executes it against Excel via the
// ExcelApp / SAFEARRAY RAII helpers in excel_com.h, and returns one
// OpResponse JSON object. Mirrors the Rust reference backend
// (languages/rust/src/ops/dispatch.rs) for params and result shapes.

#include "excel_com.h"
#include "ops.h"

#include <filesystem>
#include <string>

namespace fs = std::filesystem;
using json = nlohmann::json;

// ---------------------------------------------------------------------------
// Error helper: a C++ exception carrying a structured ExcelError payload.
// ---------------------------------------------------------------------------
struct OpError {
    std::string category;
    long long code;
    std::string message;
    std::string hint; // empty => omitted
};

static OpError make_err(const char* category, const std::string& message,
                        const std::string& hint = std::string()) {
    return OpError{category, 0, message, hint};
}

// Convert a caught _com_error into an OpError (ComError + HRESULT decimal).
static OpError com_to_err(const _com_error& e) {
    HRESULT hr = e.Error();
    std::string msg;
    // Prefer the IErrorInfo description if present.
    _bstr_t desc = e.Description();
    if (desc.length() > 0) {
        msg = static_cast<const char*>(desc);
    } else {
        const TCHAR* em = e.ErrorMessage();
        if (em) {
            // ErrorMessage returns wide string under UNICODE; convert.
            std::wstring w(em);
            msg = WStringToUtf8(w);
        } else {
            msg = "COM error";
        }
    }
    return OpError{"ComError", static_cast<long long>(static_cast<int>(hr)), msg, std::string()};
}

// ---------------------------------------------------------------------------
// Format code mapping (Excel FileFormat enum)
// ---------------------------------------------------------------------------
static int format_code(const std::string& fmt, bool* ok) {
    *ok = true;
    if (fmt == "xlsx") return 51;
    if (fmt == "xlsm") return 52;
    if (fmt == "xls")  return 56;
    if (fmt == "csv")  return 6;
    *ok = false;
    return 51;
}

// Derive a format code from a path's extension; fall back to xlsx (51).
static int format_from_ext(const std::string& path) {
    fs::path p(path);
    std::string ext = p.extension().string();
    if (!ext.empty() && ext[0] == '.') ext.erase(0, 1);
    for (auto& c : ext) c = static_cast<char>(::tolower(c));
    bool ok = false;
    int code = format_code(ext, &ok);
    return ok ? code : 51;
}

// ---------------------------------------------------------------------------
// Request field accessors
// ---------------------------------------------------------------------------
static std::string get_path(const json& req) {
    if (!req.contains("path") || !req["path"].is_string()) {
        throw make_err("InvalidArg", "request is missing required string field 'path'");
    }
    return req["path"].get<std::string>();
}

static bool has_target_range(const json& req) {
    return req.contains("target") && req["target"].is_object()
        && req["target"].contains("range") && req["target"]["range"].is_string();
}

static std::string get_target_range(const json& req, const char* op) {
    if (!has_target_range(req)) {
        throw make_err("InvalidArg", std::string(op) + " requires target.range");
    }
    return req["target"]["range"].get<std::string>();
}

// Returns the sheet name if present, else empty optional via bool flag.
static bool get_target_sheet(const json& req, std::string& out) {
    if (req.contains("target") && req["target"].is_object()
        && req["target"].contains("sheet") && req["target"]["sheet"].is_string()) {
        out = req["target"]["sheet"].get<std::string>();
        return true;
    }
    return false;
}

static const json& get_params(const json& req) {
    static const json empty = json::object();
    if (req.contains("params") && req["params"].is_object()) return req["params"];
    return empty;
}

// ---------------------------------------------------------------------------
// Sheet resolution
// ---------------------------------------------------------------------------
// Write semantics: name miss falls back to the first sheet (handles
// non-English default sheet names on freshly created workbooks).
static Excel::_WorksheetPtr resolve_sheet_write(ExcelApp& app, Excel::_WorkbookPtr& wb,
                                                const json& req) {
    std::string sheet;
    if (get_target_sheet(req, sheet)) {
        try {
            return app.GetSheet(wb, Utf8ToWString(sheet));
        } catch (const _com_error&) {
            return app.GetSheet(wb, 1);
        }
    }
    return app.GetSheet(wb, 1);
}

// Read semantics: name miss => SheetNotFound.
static Excel::_WorksheetPtr resolve_sheet_read(ExcelApp& app, Excel::_WorkbookPtr& wb,
                                               const json& req) {
    std::string sheet;
    if (get_target_sheet(req, sheet)) {
        try {
            return app.GetSheet(wb, Utf8ToWString(sheet));
        } catch (const _com_error&) {
            throw make_err("SheetNotFound", "sheet not found: " + sheet);
        }
    }
    return app.GetSheet(wb, 1);
}

// ---------------------------------------------------------------------------
// open_or_create / open
// ---------------------------------------------------------------------------
static std::wstring win_path(const std::string& path) {
    std::wstring w = Utf8ToWString(path);
    for (auto& c : w) if (c == L'/') c = L'\\';
    return w;
}

static Excel::_WorkbookPtr open_or_create(ExcelApp& app, const std::string& path) {
    if (fs::exists(fs::path(path))) {
        return app.OpenWorkbook(win_path(path));
    }
    return app.AddWorkbook();
}

static Excel::_WorkbookPtr open_existing(ExcelApp& app, const std::string& path) {
    return app.OpenWorkbook(win_path(path));
}

// ---------------------------------------------------------------------------
// save_and_close — shared by all mutating ops
// ---------------------------------------------------------------------------
static void save_and_close(Excel::_WorkbookPtr& wb, const json& req, const std::string& path) {
    std::wstring savePath;
    int fmt;
    if (req.contains("save_as") && req["save_as"].is_object()) {
        const json& sa = req["save_as"];
        if (!sa.contains("path") || !sa["path"].is_string()) {
            throw make_err("InvalidArg", "save_as requires a string 'path'");
        }
        std::string fmtStr = sa.value("format", std::string("xlsx"));
        bool ok = false;
        fmt = format_code(fmtStr, &ok);
        if (!ok) throw make_err("UnsupportedFormat", "unknown format: " + fmtStr);
        savePath = win_path(sa["path"].get<std::string>());
    } else {
        fmt = format_from_ext(path);
        savePath = win_path(path);
    }
    wb->SaveAs(_variant_t(savePath.c_str()),
               _variant_t(static_cast<long>(fmt)),
               vtMissing, vtMissing, vtMissing, vtMissing,
               Excel::xlNoChange, vtMissing, vtMissing, vtMissing,
               vtMissing, vtMissing);
    wb->Close(_variant_t(VARIANT_FALSE));
}

static void close_no_save(Excel::_WorkbookPtr& wb) {
    wb->Close(_variant_t(VARIANT_FALSE));
}

// ---------------------------------------------------------------------------
// Variant <-> JSON helpers
// ---------------------------------------------------------------------------
static json variant_to_json(const _variant_t& v) {
    switch (v.vt) {
        case VT_BSTR:
            return json{{"kind", "string"}, {"value", WStringToUtf8(VariantToWString(v))}};
        case VT_R8:
            return json{{"kind", "number"}, {"value", v.dblVal}};
        case VT_R4:
            return json{{"kind", "number"}, {"value", static_cast<double>(v.fltVal)}};
        case VT_I4:
            return json{{"kind", "number"}, {"value", static_cast<double>(v.lVal)}};
        case VT_I2:
            return json{{"kind", "number"}, {"value", static_cast<double>(v.iVal)}};
        case VT_BOOL:
            return json{{"kind", "bool"}, {"value", v.boolVal != VARIANT_FALSE}};
        case VT_EMPTY:
        case VT_NULL:
            return json{{"kind", "empty"}, {"value", nullptr}};
        default:
            return json{{"kind", "unknown"}, {"value", nullptr}};
    }
}

// Convert a JSON scalar to a _variant_t (string/number/bool/null).
static _variant_t json_scalar_to_variant(const json& v) {
    if (v.is_string()) return _variant_t(v.get<std::string>().c_str());
    if (v.is_number()) return _variant_t(v.get<double>());
    if (v.is_boolean()) return _variant_t(v.get<bool>());
    if (v.is_null()) { _variant_t e; e.vt = VT_EMPTY; return e; }
    throw make_err("InvalidArg", "params.values cells must be string, number, bool, or null");
}

// ===========================================================================
// Ops
// ===========================================================================

// cell.write -----------------------------------------------------------------
static json op_cell_write(const json& req) {
    std::string path = get_path(req);
    std::string range = get_target_range(req, "cell.write");
    const json& params = get_params(req);
    if (!params.contains("kind") || !params["kind"].is_string()) {
        throw make_err("InvalidArg", "params.kind is required");
    }
    std::string kind = params["kind"].get<std::string>();
    if (!params.contains("value")) {
        throw make_err("InvalidArg", "params.value is required");
    }
    const json& value = params["value"];

    ExcelApp app;
    Excel::_WorkbookPtr wb = open_or_create(app, path);
    Excel::_WorksheetPtr ws = resolve_sheet_write(app, wb, req);
    Excel::RangePtr cell = app.GetRange(ws, Utf8ToWString(range));

    if (kind == "string") {
        std::string s = value.is_string() ? value.get<std::string>() : value.dump();
        cell->PutValue2(_variant_t(Utf8ToWString(s).c_str()));
    } else if (kind == "number") {
        if (!value.is_number()) throw make_err("InvalidArg", "params.value must be a number for kind=number");
        cell->PutValue2(_variant_t(value.get<double>()));
    } else if (kind == "bool") {
        if (!value.is_boolean()) throw make_err("InvalidArg", "params.value must be a bool for kind=bool");
        cell->PutValue2(_variant_t(value.get<bool>()));
    } else if (kind == "formula") {
        if (!value.is_string()) throw make_err("InvalidArg", "params.value must be a string for kind=formula");
        cell->PutFormula(_variant_t(Utf8ToWString(value.get<std::string>()).c_str()));
    } else {
        throw make_err("InvalidArg", "unknown kind: " + kind);
    }

    save_and_close(wb, req, path);
    return json{{"written", true}};
}

// range.read -----------------------------------------------------------------
static json op_range_read(const json& req) {
    std::string path = get_path(req);
    std::string range = get_target_range(req, "range.read");
    const json& params = get_params(req);
    std::string property = "value2";
    if (params.contains("property") && params["property"].is_string())
        property = params["property"].get<std::string>();

    if (!fs::exists(fs::path(path)))
        throw make_err("FileNotFound", "file not found: " + path);

    ExcelApp app;
    Excel::_WorkbookPtr wb = open_existing(app, path);
    Excel::_WorksheetPtr ws = resolve_sheet_read(app, wb, req);
    Excel::RangePtr cell = app.GetRange(ws, Utf8ToWString(range));

    json result;
    if (property == "value2") {
        result = variant_to_json(cell->GetValue2());
    } else if (property == "formula") {
        _variant_t v = cell->GetFormula();
        if (v.vt == VT_BSTR) result = json{{"kind", "string"}, {"value", WStringToUtf8(VariantToWString(v))}};
        else if (v.vt == VT_EMPTY || v.vt == VT_NULL) result = json{{"kind", "empty"}, {"value", nullptr}};
        else result = variant_to_json(v);
    } else if (property == "text") {
        _variant_t v = cell->GetText();
        if (v.vt == VT_BSTR) result = json{{"kind", "string"}, {"value", WStringToUtf8(VariantToWString(v))}};
        else if (v.vt == VT_EMPTY || v.vt == VT_NULL) result = json{{"kind", "empty"}, {"value", nullptr}};
        else result = variant_to_json(v);
    } else {
        close_no_save(wb);
        throw make_err("InvalidArg", "unknown params.property: " + property + "; expected value2|formula|text");
    }

    close_no_save(wb);
    return result;
}

// range.write_bulk -----------------------------------------------------------
static json op_range_write_bulk(const json& req) {
    std::string path = get_path(req);
    std::string range = get_target_range(req, "range.write_bulk");
    if (range.find(':') == std::string::npos)
        throw make_err("InvalidArg", "range.write_bulk requires a multi-cell target.range (e.g. \"A1:B2\")");

    const json& params = get_params(req);
    if (!params.contains("values") || !params["values"].is_array())
        throw make_err("InvalidArg", "range.write_bulk requires params.values (2D array)");
    const json& values = params["values"];

    int rows = static_cast<int>(values.size());
    if (rows == 0) throw make_err("InvalidArg", "params.values must not be empty");
    if (!values[0].is_array()) throw make_err("InvalidArg", "row 0 in params.values is not an array");
    int cols = static_cast<int>(values[0].size());
    if (cols == 0) throw make_err("InvalidArg", "params.values rows must not be empty");

    // Rectangular guard
    for (int ri = 0; ri < rows; ++ri) {
        if (!values[ri].is_array())
            throw make_err("InvalidArg", "row " + std::to_string(ri) + " in params.values is not an array");
        if (static_cast<int>(values[ri].size()) != cols)
            throw make_err("InvalidArg", "row " + std::to_string(ri) + " has " +
                std::to_string(values[ri].size()) + " elements but row 0 has " +
                std::to_string(cols) + "; all rows must have the same length");
    }

    // Build 2D SAFEARRAY(VT_VARIANT)
    std::vector<std::vector<_variant_t>> data(rows, std::vector<_variant_t>(cols));
    for (int ri = 0; ri < rows; ++ri)
        for (int ci = 0; ci < cols; ++ci)
            data[ri][ci] = json_scalar_to_variant(values[ri][ci]);
    _variant_t sa = MakeSafeArray2D(rows, cols, data);

    ExcelApp app;
    Excel::_WorkbookPtr wb = open_or_create(app, path);
    Excel::_WorksheetPtr ws = resolve_sheet_write(app, wb, req);
    Excel::RangePtr rng = app.GetRange(ws, Utf8ToWString(range));
    rng->PutValue2(sa);

    save_and_close(wb, req, path);
    return json{{"written", true}, {"rows", rows}, {"cols", cols}};
}

// range.clear ----------------------------------------------------------------
static json op_range_clear(const json& req) {
    std::string path = get_path(req);
    std::string range = get_target_range(req, "range.clear");
    const json& params = get_params(req);
    std::string mode = "contents";
    if (params.contains("mode") && params["mode"].is_string())
        mode = params["mode"].get<std::string>();

    if (!fs::exists(fs::path(path)))
        throw make_err("FileNotFound", "file not found: " + path);

    ExcelApp app;
    Excel::_WorkbookPtr wb = open_existing(app, path);
    Excel::_WorksheetPtr ws = resolve_sheet_write(app, wb, req);
    Excel::RangePtr rng = app.GetRange(ws, Utf8ToWString(range));

    if (mode == "contents") {
        rng->ClearContents();
    } else if (mode == "all") {
        rng->Clear();
    } else {
        close_no_save(wb);
        throw make_err("InvalidArg", "unknown mode: " + mode + "; expected contents|all");
    }

    save_and_close(wb, req, path);
    return json{{"cleared", true}};
}

// range.merge ----------------------------------------------------------------
static json op_range_merge(const json& req) {
    std::string path = get_path(req);
    std::string range = get_target_range(req, "range.merge");
    if (range.find(':') == std::string::npos)
        throw make_err("InvalidArg", "range.merge requires a multi-cell range (e.g. A1:C3)");

    if (!fs::exists(fs::path(path)))
        throw make_err("FileNotFound", "file not found: " + path);

    ExcelApp app;
    Excel::_WorkbookPtr wb = open_existing(app, path);
    Excel::_WorksheetPtr ws = resolve_sheet_write(app, wb, req);
    Excel::RangePtr rng = app.GetRange(ws, Utf8ToWString(range));

    // Merge(Across=false) — collapse the whole range into a single merged cell.
    rng->Merge(_variant_t(false));

    save_and_close(wb, req, path);
    return json{{"merged", true}};
}

// range.unmerge --------------------------------------------------------------
static json op_range_unmerge(const json& req) {
    std::string path = get_path(req);
    std::string range = get_target_range(req, "range.unmerge");

    if (!fs::exists(fs::path(path)))
        throw make_err("FileNotFound", "file not found: " + path);

    ExcelApp app;
    Excel::_WorkbookPtr wb = open_existing(app, path);
    Excel::_WorksheetPtr ws = resolve_sheet_write(app, wb, req);
    Excel::RangePtr rng = app.GetRange(ws, Utf8ToWString(range));

    rng->UnMerge();

    save_and_close(wb, req, path);
    return json{{"unmerged", true}};
}


// range.copy_values ----------------------------------------------------------
static json op_range_copy_values(const json& req) {
    std::string path = get_path(req);
    std::string src = get_target_range(req, "range.copy_values");
    const json& params = get_params(req);
    if (!params.contains("dest") || !params["dest"].is_string())
        throw make_err("InvalidArg", "range.copy_values requires params.dest (destination address)");
    std::string dest = params["dest"].get<std::string>();

    if (!fs::exists(fs::path(path)))
        throw make_err("FileNotFound", "file not found: " + path);

    ExcelApp app;
    Excel::_WorkbookPtr wb = open_existing(app, path);
    Excel::_WorksheetPtr ws = resolve_sheet_write(app, wb, req);

    Excel::RangePtr srcRng = app.GetRange(ws, Utf8ToWString(src));
    _variant_t srcVal = srcRng->GetValue2();
    Excel::RangePtr destRng = app.GetRange(ws, Utf8ToWString(dest));

    switch (srcVal.vt) {
        case VT_BSTR:
        case VT_R8:
        case VT_R4:
        case VT_I4:
        case VT_I2:
        case VT_BOOL:
            destRng->PutValue2(srcVal);
            break;
        case VT_EMPTY:
        case VT_NULL:
            // Source empty: nothing to copy.
            break;
        default:
            close_no_save(wb);
            throw make_err("InvalidArg", "source cell has unsupported value type for copy_values");
    }

    save_and_close(wb, req, path);
    return json{{"copied", true}};
}

// inspect --------------------------------------------------------------------
static json op_inspect(const json& req) {
    std::string path = get_path(req);
    if (!fs::exists(fs::path(path)))
        throw make_err("FileNotFound", "file not found: " + path);

    ExcelApp app;
    Excel::_WorkbookPtr wb = open_existing(app, path);
    Excel::SheetsPtr sheetsColl = wb->GetWorksheets();
    long count = sheetsColl->GetCount();

    json sheets = json::array();
    for (long i = 1; i <= count; ++i) {
        Excel::_WorksheetPtr ws = app.GetSheet(wb, static_cast<int>(i));
        std::string name = WStringToUtf8(std::wstring(static_cast<const wchar_t*>(ws->GetName())));
        Excel::RangePtr used = ws->GetUsedRange();
        std::string addr = WStringToUtf8(std::wstring(static_cast<const wchar_t*>(used->GetAddress(vtMissing, vtMissing, Excel::xlA1))));
        long rows = used->GetRows()->GetCount();
        long cols = used->GetColumns()->GetCount();
        sheets.push_back(json{
            {"index", i},
            {"name", name},
            {"used_range", addr},
            {"rows", rows},
            {"cols", cols},
        });
    }

    close_no_save(wb);
    return json{{"sheets", sheets}};
}

// set_format -----------------------------------------------------------------
static json op_set_format(const json& req) {
    std::string path = get_path(req);
    std::string range = get_target_range(req, "set_format");
    const json& params = get_params(req);

    if (!fs::exists(fs::path(path)))
        throw make_err("FileNotFound", "file not found: " + path);

    ExcelApp app;
    Excel::_WorkbookPtr wb = open_existing(app, path);
    Excel::_WorksheetPtr ws = resolve_sheet_write(app, wb, req);
    Excel::RangePtr rng = app.GetRange(ws, Utf8ToWString(range));

    if (params.contains("bold") && params["bold"].is_boolean())
        rng->GetFont()->PutBold(_variant_t(params["bold"].get<bool>()));
    if (params.contains("italic") && params["italic"].is_boolean())
        rng->GetFont()->PutItalic(_variant_t(params["italic"].get<bool>()));
    if (params.contains("font_size") && params["font_size"].is_number())
        rng->GetFont()->PutSize(_variant_t(params["font_size"].get<double>()));
    if (params.contains("font_name") && params["font_name"].is_string())
        rng->GetFont()->PutName(_variant_t(Utf8ToWString(params["font_name"].get<std::string>()).c_str()));
    if (params.contains("font_color") && params["font_color"].is_number_integer())
        rng->GetFont()->PutColor(_variant_t(static_cast<long>(params["font_color"].get<long long>())));
    if (params.contains("bg_color") && params["bg_color"].is_number_integer())
        rng->GetInterior()->PutColor(_variant_t(static_cast<long>(params["bg_color"].get<long long>())));
    if (params.contains("number_format") && params["number_format"].is_string())
        rng->PutNumberFormat(_variant_t(Utf8ToWString(params["number_format"].get<std::string>()).c_str()));

    save_and_close(wb, req, path);
    return json{{"formatted", true}};
}

// row.insert -----------------------------------------------------------------
static json op_row_insert(const json& req) {
    std::string path = get_path(req);
    const json& params = get_params(req);
    if (!params.contains("row") || !params["row"].is_number_integer())
        throw make_err("InvalidArg", "row.insert requires params.row");
    long row = static_cast<long>(params["row"].get<long long>());
    long count = 1;
    if (params.contains("count") && params["count"].is_number_integer())
        count = static_cast<long>(params["count"].get<long long>());

    if (!fs::exists(fs::path(path)))
        throw make_err("FileNotFound", "file not found: " + path);

    ExcelApp app;
    Excel::_WorkbookPtr wb = open_existing(app, path);
    Excel::_WorksheetPtr ws = resolve_sheet_write(app, wb, req);

    for (long i = 0; i < count; ++i) {
        std::wstring addr = std::to_wstring(row) + L":" + std::to_wstring(row);
        Excel::RangePtr rowRng = ws->GetRange(_variant_t(addr.c_str()));
        rowRng->Insert(vtMissing, vtMissing);
    }

    save_and_close(wb, req, path);
    return json{{"inserted", true}, {"row", row}, {"count", count}};
}

// row.delete -----------------------------------------------------------------
static json op_row_delete(const json& req) {
    std::string path = get_path(req);
    const json& params = get_params(req);
    if (!params.contains("row") || !params["row"].is_number_integer())
        throw make_err("InvalidArg", "row.delete requires params.row");
    long row = static_cast<long>(params["row"].get<long long>());
    long count = 1;
    if (params.contains("count") && params["count"].is_number_integer())
        count = static_cast<long>(params["count"].get<long long>());

    if (!fs::exists(fs::path(path)))
        throw make_err("FileNotFound", "file not found: " + path);

    ExcelApp app;
    Excel::_WorkbookPtr wb = open_existing(app, path);
    Excel::_WorksheetPtr ws = resolve_sheet_write(app, wb, req);

    std::wstring addr = std::to_wstring(row) + L":" + std::to_wstring(row + count - 1);
    Excel::RangePtr rowRng = ws->GetRange(_variant_t(addr.c_str()));
    rowRng->Delete(vtMissing);

    save_and_close(wb, req, path);
    return json{{"deleted", true}, {"row", row}, {"count", count}};
}

// sheet.add ------------------------------------------------------------------
static json op_sheet_add(const json& req) {
    std::string path = get_path(req);
    const json& params = get_params(req);

    ExcelApp app;
    Excel::_WorkbookPtr wb = open_or_create(app, path);
    Excel::_WorksheetPtr ws = app.AddSheet(wb);
    if (params.contains("name") && params["name"].is_string())
        ws->PutName(_bstr_t(Utf8ToWString(params["name"].get<std::string>()).c_str()));
    std::string actual = WStringToUtf8(std::wstring(static_cast<const wchar_t*>(ws->GetName())));

    save_and_close(wb, req, path);
    return json{{"added", true}, {"name", actual}};
}

// sheet.rename ---------------------------------------------------------------
static json op_sheet_rename(const json& req) {
    std::string path = get_path(req);
    std::string sheet;
    if (!get_target_sheet(req, sheet))
        throw make_err("InvalidArg", "sheet.rename requires target.sheet");
    const json& params = get_params(req);
    if (!params.contains("name") || !params["name"].is_string())
        throw make_err("InvalidArg", "sheet.rename requires params.name");
    std::string newName = params["name"].get<std::string>();

    if (!fs::exists(fs::path(path)))
        throw make_err("FileNotFound", "file not found: " + path);

    ExcelApp app;
    Excel::_WorkbookPtr wb = open_existing(app, path);
    Excel::_WorksheetPtr ws;
    try {
        ws = app.GetSheet(wb, Utf8ToWString(sheet));
    } catch (const _com_error&) {
        close_no_save(wb);
        throw make_err("SheetNotFound", "sheet not found: " + sheet);
    }
    ws->PutName(_bstr_t(Utf8ToWString(newName).c_str()));

    save_and_close(wb, req, path);
    return json{{"renamed", true}, {"old_name", sheet}, {"new_name", newName}};
}

// sheet.delete ---------------------------------------------------------------
static json op_sheet_delete(const json& req) {
    std::string path = get_path(req);
    std::string sheet;
    if (!get_target_sheet(req, sheet))
        throw make_err("InvalidArg", "sheet.delete requires target.sheet");

    if (!fs::exists(fs::path(path)))
        throw make_err("FileNotFound", "file not found: " + path);

    ExcelApp app;
    Excel::_WorkbookPtr wb = open_existing(app, path);
    Excel::_WorksheetPtr ws;
    try {
        ws = app.GetSheet(wb, Utf8ToWString(sheet));
    } catch (const _com_error&) {
        close_no_save(wb);
        throw make_err("SheetNotFound", "sheet not found: " + sheet);
    }
    ws->Delete();

    save_and_close(wb, req, path);
    return json{{"deleted", true}, {"name", sheet}};
}

// ===========================================================================
// Dispatch
// ===========================================================================
static json build_err(const OpError& e) {
    json err = {
        {"category", e.category},
        {"code", e.code},
        {"message", e.message},
    };
    if (!e.hint.empty()) err["hint"] = e.hint;
    return json{{"ok", false}, {"error", err}};
}

json dispatch_op(const json& req) {
    std::string op;
    if (req.contains("op") && req["op"].is_string()) op = req["op"].get<std::string>();
    else return build_err(make_err("InvalidArg", "request is missing required string field 'op'"));

    try {
        json result;
        if      (op == "cell.write")        result = op_cell_write(req);
        else if (op == "range.read")        result = op_range_read(req);
        else if (op == "range.write_bulk")  result = op_range_write_bulk(req);
        else if (op == "range.clear")       result = op_range_clear(req);
        else if (op == "range.merge")        result = op_range_merge(req);
        else if (op == "range.unmerge")      result = op_range_unmerge(req);
        else if (op == "range.copy_values") result = op_range_copy_values(req);
        else if (op == "inspect")           result = op_inspect(req);
        else if (op == "set_format")        result = op_set_format(req);
        else if (op == "row.insert")        result = op_row_insert(req);
        else if (op == "row.delete")        result = op_row_delete(req);
        else if (op == "sheet.add")         result = op_sheet_add(req);
        else if (op == "sheet.rename")      result = op_sheet_rename(req);
        else if (op == "sheet.delete")      result = op_sheet_delete(req);
        else {
            return build_err(make_err("Unknown", "unknown op: " + op,
                "supported: cell.write, range.read, range.write_bulk, range.clear, range.merge, range.unmerge, range.copy_values, "
                "inspect, set_format, row.insert, row.delete, sheet.add, sheet.rename, sheet.delete"));
        }
        return json{{"ok", true}, {"result", result}};
    } catch (const OpError& e) {
        return build_err(e);
    } catch (const _com_error& e) {
        return build_err(com_to_err(e));
    } catch (const std::exception& e) {
        return build_err(OpError{"Unknown", 0, std::string("internal error: ") + e.what(), std::string()});
    } catch (...) {
        return build_err(OpError{"Unknown", 0, "unknown internal error", std::string()});
    }
}
