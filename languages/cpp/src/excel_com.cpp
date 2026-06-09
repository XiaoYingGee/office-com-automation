// excel_com.cpp - Implementation of RAII wrappers, SAFEARRAY helpers, and COM utilities

#include "excel_com.h"
#include <iostream>

// ---------------------------------------------------------------------------
// ExcelApp
// ---------------------------------------------------------------------------

ExcelApp::ExcelApp() {
    HRESULT hr = app_.CreateInstance(__uuidof(Excel::Application));
    if (FAILED(hr)) {
        throw std::runtime_error("Failed to create Excel.Application instance");
    }
    app_->PutVisible(VARIANT_FALSE);
    app_->PutDisplayAlerts(VARIANT_FALSE);
    app_->PutScreenUpdating(VARIANT_FALSE);
}

ExcelApp::~ExcelApp() {
    try {
        if (app_) {
            app_->PutScreenUpdating(VARIANT_TRUE);
            app_->PutDisplayAlerts(VARIANT_TRUE);
            app_->Quit();
        }
    } catch (...) {
        // Swallow errors during cleanup
    }
    app_ = nullptr;  // Release COM reference
}

Excel::_WorkbookPtr ExcelApp::AddWorkbook() {
    return app_->GetWorkbooks()->Add(vtMissing);
}

Excel::_WorkbookPtr ExcelApp::OpenWorkbook(const std::wstring& path) {
    return app_->GetWorkbooks()->Open(_bstr_t(path.c_str()));
}

Excel::_WorksheetPtr ExcelApp::GetSheet(Excel::_WorkbookPtr& wb, int index) {
    return wb->GetWorksheets()->GetItem(_variant_t(static_cast<long>(index)));
}

Excel::_WorksheetPtr ExcelApp::GetSheet(Excel::_WorkbookPtr& wb, const std::wstring& name) {
    return wb->GetWorksheets()->GetItem(_variant_t(name.c_str()));
}

Excel::_WorksheetPtr ExcelApp::AddSheet(Excel::_WorkbookPtr& wb) {
    return wb->GetWorksheets()->Add(vtMissing, vtMissing, vtMissing, vtMissing);
}

Excel::_WorksheetPtr ExcelApp::AddSheetAfter(Excel::_WorkbookPtr& wb, Excel::_WorksheetPtr& afterSheet) {
    return wb->GetWorksheets()->Add(vtMissing, _variant_t(static_cast<IDispatch*>(afterSheet)), vtMissing, vtMissing);
}

Excel::RangePtr ExcelApp::GetRange(Excel::_WorksheetPtr& ws, const std::wstring& address) {
    return ws->GetRange(_variant_t(address.c_str()));
}

Excel::RangePtr ExcelApp::GetCell(Excel::_WorksheetPtr& ws, int row, int col) {
    std::wstring addr = CellAddress(row, col);
    return ws->GetRange(_variant_t(addr.c_str()));
}

void ExcelApp::BulkWrite(Excel::_WorksheetPtr& ws, int startRow, int startCol,
                          int rows, int cols, const _variant_t& safeArrayVar) {
    std::wstring addr = RangeAddress(startRow, startCol, startRow + rows - 1, startCol + cols - 1);
    Excel::RangePtr rng = ws->GetRange(_variant_t(addr.c_str()));
    rng->PutValue2(safeArrayVar);
}

// ---------------------------------------------------------------------------
// SAFEARRAY helpers
// ---------------------------------------------------------------------------

_variant_t MakeSafeArray2DSequential(int rows, int cols) {
    // SAFEARRAY bounds: dimension 0 = rows (1-based), dimension 1 = cols (1-based)
    SAFEARRAYBOUND bounds[2];
    bounds[0].lLbound = 1;
    bounds[0].cElements = static_cast<ULONG>(rows);
    bounds[1].lLbound = 1;
    bounds[1].cElements = static_cast<ULONG>(cols);

    SAFEARRAY* psa = ::SafeArrayCreate(VT_VARIANT, 2, bounds);
    if (!psa) throw std::runtime_error("SafeArrayCreate failed");

    for (int r = 1; r <= rows; ++r) {
        for (int c = 1; c <= cols; ++c) {
            LONG indices[2] = { r, c };
            _variant_t val(static_cast<double>((r - 1) * cols + c));
            HRESULT hr = ::SafeArrayPutElement(psa, indices, &val);
            if (FAILED(hr)) {
                ::SafeArrayDestroy(psa);
                throw std::runtime_error("SafeArrayPutElement failed");
            }
        }
    }

    _variant_t result;
    result.vt = VT_ARRAY | VT_VARIANT;
    result.parray = psa;
    return result;
}

_variant_t MakeSafeArray2D(int rows, int cols, const std::vector<std::vector<_variant_t>>& data) {
    SAFEARRAYBOUND bounds[2];
    bounds[0].lLbound = 1;
    bounds[0].cElements = static_cast<ULONG>(rows);
    bounds[1].lLbound = 1;
    bounds[1].cElements = static_cast<ULONG>(cols);

    SAFEARRAY* psa = ::SafeArrayCreate(VT_VARIANT, 2, bounds);
    if (!psa) throw std::runtime_error("SafeArrayCreate failed");

    for (int r = 0; r < rows; ++r) {
        for (int c = 0; c < cols; ++c) {
            LONG indices[2] = { r + 1, c + 1 };
            VARIANT v;
            ::VariantInit(&v);
            ::VariantCopy(&v, const_cast<VARIANT*>(static_cast<const VARIANT*>(&data[r][c])));
            HRESULT hr = ::SafeArrayPutElement(psa, indices, &v);
            ::VariantClear(&v);
            if (FAILED(hr)) {
                ::SafeArrayDestroy(psa);
                throw std::runtime_error("SafeArrayPutElement failed");
            }
        }
    }

    _variant_t result;
    result.vt = VT_ARRAY | VT_VARIANT;
    result.parray = psa;
    return result;
}

// ---------------------------------------------------------------------------
// COM variant utilities
// ---------------------------------------------------------------------------

void SaveAsXlsx(Excel::_WorkbookPtr& wb, const std::wstring& path) {
    wb->SaveAs(_variant_t(path.c_str()),
               _variant_t(static_cast<long>(Excel::xlOpenXMLWorkbook)),
               vtMissing, vtMissing, vtMissing, vtMissing,
               Excel::xlNoChange, vtMissing, vtMissing, vtMissing,
               vtMissing, vtMissing);
}

std::wstring VariantToWString(const _variant_t& v) {
    if (v.vt == VT_BSTR && v.bstrVal) {
        return std::wstring(static_cast<const wchar_t*>(_bstr_t(v.bstrVal)));
    }
    return L"";
}

double VariantToDouble(const _variant_t& v, double defaultVal) {
    if (v.vt == VT_R8) return v.dblVal;
    if (v.vt == VT_I4) return static_cast<double>(v.lVal);
    if (v.vt == VT_I2) return static_cast<double>(v.iVal);
    return defaultVal;
}

// ---------------------------------------------------------------------------
// UTF-8 <-> UTF-16 conversion (Win32 MultiByteToWideChar / WideCharToMultiByte)
// ---------------------------------------------------------------------------
std::wstring Utf8ToWString(const std::string& s) {
    if (s.empty()) return std::wstring();
    int len = ::MultiByteToWideChar(CP_UTF8, 0, s.data(), static_cast<int>(s.size()), nullptr, 0);
    std::wstring w(len, L'\0');
    ::MultiByteToWideChar(CP_UTF8, 0, s.data(), static_cast<int>(s.size()), &w[0], len);
    return w;
}

std::string WStringToUtf8(const std::wstring& w) {
    if (w.empty()) return std::string();
    int len = ::WideCharToMultiByte(CP_UTF8, 0, w.data(), static_cast<int>(w.size()), nullptr, 0, nullptr, nullptr);
    std::string s(len, '\0');
    ::WideCharToMultiByte(CP_UTF8, 0, w.data(), static_cast<int>(w.size()), &s[0], len, nullptr, nullptr);
    return s;
}
