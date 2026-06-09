#pragma once
// excel_com.h - RAII wrappers for Excel COM automation via #import + smart pointers
//
// Provides:
//   CoInitGuard  - RAII wrapper for CoInitializeEx / CoUninitialize
//   ExcelApp     - RAII wrapper for Excel.Application lifecycle
//   MakeSafeArray2D - Helper to build 2D SAFEARRAY(VT_VARIANT)

#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <windows.h>
#include <comdef.h>
#include <objbase.h>

#include <string>
#include <vector>
#include <stdexcept>

// Import Office (MSO) type library first - provides shared Office types.
// Stays in its natural "Office" namespace.
#import "C:\\Program Files\\Microsoft Office\\root\\VFS\\ProgramFilesCommonX64\\Microsoft Shared\\OFFICE16\\MSO.DLL" \
    rename("RGB", "MsoRGB") \
    rename("DocumentProperties", "MsoDocumentProperties") \
    rename("SearchPath", "MsoSearchPath") \
    no_dual_interfaces

// Import VBE type library - stays in "VBIDE" namespace.
#import "C:\\Program Files\\Microsoft Office\\root\\VFS\\ProgramFilesCommonX86\\Microsoft Shared\\VBA\\VBA6\\VBE6EXT.OLB" \
    no_dual_interfaces

// NOTE: These using-directives are required here (not in .cpp) because the
// EXCEL.EXE #import below generates a .tlh that cross-references Office and
// VBIDE types.  Without these, the generated header will not compile.
using namespace Office;
using namespace VBIDE;

// Import Excel type library - generates Excel namespace with smart pointer types.
#import "C:\\Program Files\\Microsoft Office\\root\\Office16\\EXCEL.EXE" \
    rename("DialogBox", "ExcelDialogBox") \
    rename("RGB", "ExcelRGB") \
    rename("CopyFile", "ExcelCopyFile") \
    rename("ReplaceText", "ExcelReplaceText") \
    exclude("IFont", "IPicture") \
    no_dual_interfaces

// ---------------------------------------------------------------------------
// Cell address helper: convert 1-based (row, col) to "A1" style address
// ---------------------------------------------------------------------------
inline std::wstring CellAddress(int row, int col) {
    std::wstring colStr;
    int c = col;
    while (c > 0) {
        int rem = (c - 1) % 26;
        colStr = static_cast<wchar_t>(L'A' + rem) + colStr;
        c = (c - 1) / 26;
    }
    return colStr + std::to_wstring(row);
}

// Build a range address like "A1:J1000"
inline std::wstring RangeAddress(int r1, int c1, int r2, int c2) {
    return CellAddress(r1, c1) + L":" + CellAddress(r2, c2);
}

// ---------------------------------------------------------------------------
// CoInitGuard - RAII for COM initialization
// ---------------------------------------------------------------------------
struct CoInitGuard {
    CoInitGuard() {
        HRESULT hr = ::CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);
        if (FAILED(hr) && hr != S_FALSE) {
            throw std::runtime_error("CoInitializeEx failed");
        }
    }
    ~CoInitGuard() {
        ::CoUninitialize();
    }
    CoInitGuard(const CoInitGuard&) = delete;
    CoInitGuard& operator=(const CoInitGuard&) = delete;
};

// ---------------------------------------------------------------------------
// ExcelApp - RAII wrapper for Excel.Application
// ---------------------------------------------------------------------------
class ExcelApp {
    Excel::_ApplicationPtr app_;
public:
    ExcelApp();
    ~ExcelApp();

    ExcelApp(const ExcelApp&) = delete;
    ExcelApp& operator=(const ExcelApp&) = delete;

    Excel::_ApplicationPtr app() const { return app_; }

    // Convenience accessors
    Excel::_WorkbookPtr  AddWorkbook();
    Excel::_WorkbookPtr  OpenWorkbook(const std::wstring& path);
    Excel::_WorksheetPtr GetSheet(Excel::_WorkbookPtr& wb, int index);
    Excel::_WorksheetPtr GetSheet(Excel::_WorkbookPtr& wb, const std::wstring& name);
    Excel::_WorksheetPtr AddSheet(Excel::_WorkbookPtr& wb);
    Excel::_WorksheetPtr AddSheetAfter(Excel::_WorkbookPtr& wb, Excel::_WorksheetPtr& afterSheet);
    Excel::RangePtr      GetRange(Excel::_WorksheetPtr& ws, const std::wstring& address);
    Excel::RangePtr      GetCell(Excel::_WorksheetPtr& ws, int row, int col);

    // Bulk write: assign a 2D SAFEARRAY to a range in one COM call
    void BulkWrite(Excel::_WorksheetPtr& ws, int startRow, int startCol,
                   int rows, int cols, const _variant_t& safeArrayVar);
};

// ---------------------------------------------------------------------------
// SAFEARRAY helpers
// ---------------------------------------------------------------------------

// Build a 2D SAFEARRAY(VT_VARIANT) filled with sequential doubles: (r*cols + c + 1.0)
_variant_t MakeSafeArray2DSequential(int rows, int cols);

// Build a 2D SAFEARRAY(VT_VARIANT) from a vector of vector of _variant_t
_variant_t MakeSafeArray2D(int rows, int cols, const std::vector<std::vector<_variant_t>>& data);

// ---------------------------------------------------------------------------
// COM variant utilities
// ---------------------------------------------------------------------------

// SaveAs helper - saves workbook as .xlsx (xlOpenXMLWorkbook)
void SaveAsXlsx(Excel::_WorkbookPtr& wb, const std::wstring& path);

// Extract wstring from _variant_t with VT_BSTR
std::wstring VariantToWString(const _variant_t& v);

// Extract double from _variant_t (handles VT_R8, VT_I4, VT_I2)
double VariantToDouble(const _variant_t& v, double defaultVal = -1.0);

// UTF-8 <-> UTF-16 conversion (for JSON <-> COM string marshalling)
std::wstring Utf8ToWString(const std::string& s);
std::string  WStringToUtf8(const std::wstring& w);

// ---------------------------------------------------------------------------
// High-resolution timer using QueryPerformanceCounter
// ---------------------------------------------------------------------------
class HiResTimer {
    LARGE_INTEGER freq_, start_;
public:
    HiResTimer() {
        ::QueryPerformanceFrequency(&freq_);
        ::QueryPerformanceCounter(&start_);
    }
    void reset() {
        ::QueryPerformanceCounter(&start_);
    }
    double elapsedMs() const {
        LARGE_INTEGER now;
        ::QueryPerformanceCounter(&now);
        return static_cast<double>(now.QuadPart - start_.QuadPart) * 1000.0
             / static_cast<double>(freq_.QuadPart);
    }
};
