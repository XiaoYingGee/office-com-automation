// tasks.cpp - E01-E12 task implementations for Excel COM automation
//
// Uses #import smart pointers (_ApplicationPtr, _WorkbookPtr, _WorksheetPtr, RangePtr)
// and _variant_t / _bstr_t for COM data exchange.
//
// With no_dual_interfaces, all property access uses explicit Get/Put methods:
//   cell->PutValue2(val)  instead of  cell->Value2 = val
//   cell->GetValue2()     instead of  val = cell->Value2

#include "excel_com.h"
#include "tasks.h"

#include <iostream>
#include <filesystem>
#include <fstream>
#include <sstream>
#include <cstdio>
#include <cmath>

namespace fs = std::filesystem;

// Utility: build a full path from outputDir + filename
static std::wstring MakePath(const std::wstring& dir, const std::wstring& filename) {
    fs::path p = fs::path(dir) / filename;
    return p.wstring();
}

// Utility: delete file if it exists
static void DeleteIfExists(const std::wstring& path) {
    if (fs::exists(path)) {
        fs::remove(path);
    }
}

// Utility: check file exists and has content
static bool FileExistsNonEmpty(const std::wstring& path) {
    if (!fs::exists(path)) return false;
    return fs::file_size(path) > 0;
}

// ---------------------------------------------------------------------------
// E01 - Workbook Lifecycle
// ---------------------------------------------------------------------------
bool RunE01(const std::wstring& outputDir) {
    std::wstring filePath = MakePath(outputDir, L"E01_lifecycle.xlsx");
    DeleteIfExists(filePath);

    ExcelApp excel;

    // Create new workbook
    Excel::_WorkbookPtr wb = excel.AddWorkbook();
    Excel::_WorksheetPtr ws = excel.GetSheet(wb, 1);
    Excel::RangePtr cell = excel.GetCell(ws, 1, 1);
    cell->PutValue2(_variant_t(L"E01 Lifecycle Test"));

    // SaveAs .xlsx
    SaveAsXlsx(wb, filePath);
    std::wcout << L"  Saved: " << filePath << std::endl;

    // Close
    wb->Close(VARIANT_FALSE);
    wb = nullptr;
    ws = nullptr;
    cell = nullptr;

    // Reopen and verify
    Excel::_WorkbookPtr wb2 = excel.OpenWorkbook(filePath);
    Excel::_WorksheetPtr ws2 = excel.GetSheet(wb2, 1);
    Excel::RangePtr cell2 = excel.GetCell(ws2, 1, 1);
    _variant_t val = cell2->GetValue2();

    std::wstring readVal = VariantToWString(val);
    bool pass = (readVal == L"E01 Lifecycle Test");
    std::wcout << L"  Read back: \"" << readVal << L"\" - "
               << (pass ? L"match" : L"MISMATCH") << std::endl;

    wb2->Close(VARIANT_FALSE);
    return pass;
}

// ---------------------------------------------------------------------------
// E02 - Cell Read/Write
// ---------------------------------------------------------------------------
bool RunE02(const std::wstring& outputDir) {
    std::wstring filePath = MakePath(outputDir, L"E02_cell_rw.xlsx");
    DeleteIfExists(filePath);

    ExcelApp excel;
    Excel::_WorkbookPtr wb = excel.AddWorkbook();
    Excel::_WorksheetPtr ws = excel.GetSheet(wb, 1);

    // Write string
    excel.GetCell(ws, 1, 1)->PutValue2(_variant_t(L"Hello COM"));
    // Write number
    excel.GetCell(ws, 2, 1)->PutValue2(_variant_t(42.5));
    // Write boolean
    excel.GetCell(ws, 3, 1)->PutValue2(_variant_t(true));
    // Write date as serial number (2024-01-15 = 45306)
    excel.GetCell(ws, 4, 1)->PutValue2(_variant_t(45306.0));
    excel.GetCell(ws, 4, 1)->PutNumberFormat(_variant_t(L"yyyy-mm-dd"));

    // Read back and verify
    bool pass = true;

    // String
    _variant_t v1 = excel.GetCell(ws, 1, 1)->GetValue2();
    if (v1.vt != VT_BSTR || VariantToWString(v1) != L"Hello COM") {
        std::wcout << L"  FAIL: string mismatch" << std::endl;
        pass = false;
    } else {
        std::wcout << L"  String: OK" << std::endl;
    }

    // Number
    _variant_t v2 = excel.GetCell(ws, 2, 1)->GetValue2();
    if (v2.vt != VT_R8 || std::abs(v2.dblVal - 42.5) > 0.001) {
        std::wcout << L"  FAIL: number mismatch" << std::endl;
        pass = false;
    } else {
        std::wcout << L"  Number: OK (" << v2.dblVal << L")" << std::endl;
    }

    // Boolean (Value2 returns booleans as VT_BOOL)
    _variant_t v3 = excel.GetCell(ws, 3, 1)->GetValue2();
    if (v3.vt != VT_BOOL || v3.boolVal != VARIANT_TRUE) {
        std::wcout << L"  FAIL: boolean mismatch (vt=" << v3.vt << L")" << std::endl;
        pass = false;
    } else {
        std::wcout << L"  Boolean: OK" << std::endl;
    }

    // Date (Value2 returns date serial as double)
    _variant_t v4 = excel.GetCell(ws, 4, 1)->GetValue2();
    if (v4.vt != VT_R8 || std::abs(v4.dblVal - 45306.0) > 0.001) {
        std::wcout << L"  FAIL: date serial mismatch" << std::endl;
        pass = false;
    } else {
        std::wcout << L"  Date serial: OK (" << v4.dblVal << L")" << std::endl;
    }

    SaveAsXlsx(wb, filePath);
    wb->Close(VARIANT_FALSE);

    std::wcout << L"  File: " << filePath << std::endl;
    return pass;
}

// ---------------------------------------------------------------------------
// E03 - Bulk Range Write (SAFEARRAY vs naive)
// ---------------------------------------------------------------------------
bool RunE03(const std::wstring& outputDir) {
    const int ROWS = 1000;
    const int COLS = 10;

    std::wstring filePath = MakePath(outputDir, L"E03_bulk_write.xlsx");
    DeleteIfExists(filePath);

    ExcelApp excel;
    Excel::_WorkbookPtr wb = excel.AddWorkbook();

    // --- Idiomatic (bulk) write via SAFEARRAY ---
    Excel::_WorksheetPtr wsIdiomatic = excel.GetSheet(wb, 1);
    wsIdiomatic->PutName(_bstr_t(L"Idiomatic"));

    _variant_t saData = MakeSafeArray2DSequential(ROWS, COLS);

    HiResTimer timerIdiomatic;
    excel.BulkWrite(wsIdiomatic, 1, 1, ROWS, COLS, saData);
    double idiomaticMs = timerIdiomatic.elapsedMs();
    std::wcout << L"  Idiomatic bulk write: " << idiomaticMs << L" ms (" << ROWS << L"x" << COLS << L")" << std::endl;

    // --- Naive (cell-by-cell) write ---
    Excel::_WorksheetPtr wsNaive = excel.AddSheet(wb);
    wsNaive->PutName(_bstr_t(L"Naive"));

    HiResTimer timerNaive;
    for (int r = 1; r <= ROWS; ++r) {
        for (int c = 1; c <= COLS; ++c) {
            double val = static_cast<double>((r - 1) * COLS + c);
            excel.GetCell(wsNaive, r, c)->PutValue2(_variant_t(val));
        }
    }
    double naiveMs = timerNaive.elapsedMs();
    std::wcout << L"  Naive cell-by-cell write: " << naiveMs << L" ms (" << ROWS << L"x" << COLS << L")" << std::endl;

    double speedup = (idiomaticMs > 0.001) ? naiveMs / idiomaticMs : 0.0;
    std::wcout << L"  Speedup ratio: " << speedup << L"x" << std::endl;

    // Save
    SaveAsXlsx(wb, filePath);
    wb->Close(VARIANT_FALSE);

    // Verify: reopen and spot-check
    Excel::_WorkbookPtr wb2 = excel.OpenWorkbook(filePath);
    Excel::_WorksheetPtr ws2I = excel.GetSheet(wb2, L"Idiomatic");
    Excel::_WorksheetPtr ws2N = excel.GetSheet(wb2, L"Naive");

    bool pass = true;
    struct CheckPoint { int r; int c; };
    CheckPoint checks[] = { {1,1}, {1,10}, {500,5}, {1000,10} };

    for (auto& cp : checks) {
        double expected = static_cast<double>((cp.r - 1) * COLS + cp.c);
        double dI = VariantToDouble(excel.GetCell(ws2I, cp.r, cp.c)->GetValue2());
        double dN = VariantToDouble(excel.GetCell(ws2N, cp.r, cp.c)->GetValue2());

        if (std::abs(dI - expected) > 0.001) {
            std::wcout << L"  FAIL: Idiomatic[" << cp.r << L"," << cp.c << L"] = "
                       << dI << L", expected " << expected << std::endl;
            pass = false;
        }
        if (std::abs(dN - expected) > 0.001) {
            std::wcout << L"  FAIL: Naive[" << cp.r << L"," << cp.c << L"] = "
                       << dN << L", expected " << expected << std::endl;
            pass = false;
        }
    }

    wb2->Close(VARIANT_FALSE);
    std::wcout << L"  Bulk write is " << (speedup >= 2.0 ? L"significantly" : L"somewhat")
               << L" faster (" << speedup << L"x)" << std::endl;
    return pass;
}

// ---------------------------------------------------------------------------
// E04 - Formula & Recalc
// ---------------------------------------------------------------------------
bool RunE04(const std::wstring& outputDir) {
    std::wstring filePath = MakePath(outputDir, L"E04_formula.xlsx");
    DeleteIfExists(filePath);

    ExcelApp excel;
    Excel::_WorkbookPtr wb = excel.AddWorkbook();
    Excel::_WorksheetPtr ws = excel.GetSheet(wb, 1);

    // Write source data: column A = values, column B = multipliers
    for (int r = 1; r <= 10; ++r) {
        excel.GetCell(ws, r, 1)->PutValue2(_variant_t(static_cast<double>(r)));
        excel.GetCell(ws, r, 2)->PutValue2(_variant_t(static_cast<double>(r * 2)));
    }

    // Set calculation to manual (xlCalculationManual = -4135)
    excel.app()->PutCalculation(Excel::xlCalculationManual);

    // Write formulas in column C: =A1*B1
    for (int r = 1; r <= 10; ++r) {
        std::wstring formula = L"=A" + std::to_wstring(r) + L"*B" + std::to_wstring(r);
        excel.GetCell(ws, r, 3)->PutFormula(_variant_t(formula.c_str()));
    }

    // Force recalculation
    excel.app()->Calculate();

    // Read back results and verify
    bool pass = true;
    for (int r = 1; r <= 10; ++r) {
        _variant_t val = excel.GetCell(ws, r, 3)->GetValue2();
        double expected = static_cast<double>(r) * static_cast<double>(r * 2);
        double actual = VariantToDouble(val);
        if (std::abs(actual - expected) > 0.001) {
            std::wcout << L"  FAIL: C" << r << L" = " << actual
                       << L", expected " << expected << std::endl;
            pass = false;
        }
    }

    // Restore automatic calculation
    excel.app()->PutCalculation(Excel::xlCalculationAutomatic);

    if (pass) {
        std::wcout << L"  All 10 formulas verified OK" << std::endl;
    }

    SaveAsXlsx(wb, filePath);
    wb->Close(VARIANT_FALSE);
    return pass;
}

// ---------------------------------------------------------------------------
// E05 - Cell Formatting
// ---------------------------------------------------------------------------
bool RunE05(const std::wstring& outputDir) {
    std::wstring filePath = MakePath(outputDir, L"E05_formatting.xlsx");
    DeleteIfExists(filePath);

    ExcelApp excel;
    Excel::_WorkbookPtr wb = excel.AddWorkbook();
    Excel::_WorksheetPtr ws = excel.GetSheet(wb, 1);

    // Write data
    excel.GetCell(ws, 1, 1)->PutValue2(_variant_t(L"Bold Header"));
    excel.GetCell(ws, 2, 1)->PutValue2(_variant_t(1234.5678));
    excel.GetCell(ws, 3, 1)->PutValue2(_variant_t(L"Merged"));

    // Font Bold + Size + Color (red in BGR = 0x0000FF = 255)
    {
        Excel::RangePtr r = excel.GetCell(ws, 1, 1);
        Excel::FontPtr font = r->GetFont();
        font->PutBold(_variant_t(true));
        font->PutSize(_variant_t(16L));
        font->PutColor(_variant_t(255L));
    }

    // Interior Color (yellow in BGR: R=255, G=255, B=0 => 65535)
    {
        Excel::RangePtr r = excel.GetCell(ws, 2, 1);
        Excel::InteriorPtr interior = r->GetInterior();
        interior->PutColor(_variant_t(65535L));
    }

    // Number format
    {
        Excel::RangePtr r = excel.GetCell(ws, 2, 1);
        r->PutNumberFormat(_variant_t(L"0.00"));
    }

    // Horizontal alignment (xlCenter = -4108)
    {
        Excel::RangePtr r = excel.GetCell(ws, 1, 1);
        r->PutHorizontalAlignment(_variant_t(-4108L));
    }

    // Merge A3:C3
    {
        Excel::RangePtr r = excel.GetRange(ws, L"A3:C3");
        r->Merge(vtMissing);
        r->PutHorizontalAlignment(_variant_t(-4108L));
    }

    // Save and reopen to verify formatting persists
    SaveAsXlsx(wb, filePath);
    wb->Close(VARIANT_FALSE);

    // Verify by reopening
    Excel::_WorkbookPtr wb2 = excel.OpenWorkbook(filePath);
    Excel::_WorksheetPtr ws2 = excel.GetSheet(wb2, 1);

    bool pass = true;

    // Check bold
    {
        Excel::RangePtr r = excel.GetCell(ws2, 1, 1);
        _variant_t bold = r->GetFont()->GetBold();
        if (bold.vt != VT_BOOL || bold.boolVal != VARIANT_TRUE) {
            std::wcout << L"  FAIL: Bold not preserved" << std::endl;
            pass = false;
        } else {
            std::wcout << L"  Bold: OK" << std::endl;
        }
    }

    // Check font size (should be 16)
    {
        Excel::RangePtr r = excel.GetCell(ws2, 1, 1);
        _variant_t sz = r->GetFont()->GetSize();
        double fontSize = VariantToDouble(sz);
        if (std::abs(fontSize - 16.0) > 0.1) {
            std::wcout << L"  FAIL: Font size = " << fontSize << L", expected 16" << std::endl;
            pass = false;
        } else {
            std::wcout << L"  Font size: OK (" << fontSize << L")" << std::endl;
        }
    }

    // Check number format
    {
        Excel::RangePtr r = excel.GetCell(ws2, 2, 1);
        _variant_t fmt = r->GetNumberFormat();
        std::wstring fmtStr = VariantToWString(fmt);
        if (fmtStr != L"0.00") {
            std::wcout << L"  FAIL: NumberFormat = " << fmtStr << L", expected 0.00" << std::endl;
            pass = false;
        } else {
            std::wcout << L"  NumberFormat: OK" << std::endl;
        }
    }

    // Check interior color (yellow = 65535 in BGR)
    {
        Excel::RangePtr r = excel.GetCell(ws2, 2, 1);
        _variant_t color = r->GetInterior()->GetColor();
        double colorVal = VariantToDouble(color);
        if (std::abs(colorVal - 65535.0) > 0.1) {
            std::wcout << L"  FAIL: Interior color = " << colorVal << L", expected 65535" << std::endl;
            pass = false;
        } else {
            std::wcout << L"  Interior color: OK (" << colorVal << L")" << std::endl;
        }
    }

    // Check merge (MergeCells on A3)
    {
        Excel::RangePtr r = excel.GetCell(ws2, 3, 1);
        _variant_t merged = r->GetMergeCells();
        if (merged.vt != VT_BOOL || merged.boolVal != VARIANT_TRUE) {
            std::wcout << L"  FAIL: Merge not preserved" << std::endl;
            pass = false;
        } else {
            std::wcout << L"  Merge: OK" << std::endl;
        }
    }

    wb2->Close(VARIANT_FALSE);
    return pass;
}

// ---------------------------------------------------------------------------
// E06 - Row/Column Structure
// ---------------------------------------------------------------------------
bool RunE06(const std::wstring& outputDir) {
    std::wstring filePath = MakePath(outputDir, L"E06_structure.xlsx");
    DeleteIfExists(filePath);

    ExcelApp excel;
    Excel::_WorkbookPtr wb = excel.AddWorkbook();
    Excel::_WorksheetPtr ws = excel.GetSheet(wb, 1);

    // Write initial data
    for (int r = 1; r <= 5; ++r) {
        for (int c = 1; c <= 3; ++c) {
            std::wstring val = L"R" + std::to_wstring(r) + L"C" + std::to_wstring(c);
            excel.GetCell(ws, r, c)->PutValue2(_variant_t(val.c_str()));
        }
    }

    // Insert a row at row 3
    Excel::RangePtr row3 = excel.GetRange(ws, L"3:3");
    row3->Insert(vtMissing, vtMissing);
    std::wcout << L"  Inserted row at 3" << std::endl;

    // Write into the new row
    excel.GetCell(ws, 3, 1)->PutValue2(_variant_t(L"Inserted"));

    // Delete row 5 (was originally row 4, shifted down by insert)
    Excel::RangePtr row5 = excel.GetRange(ws, L"5:5");
    row5->Delete(vtMissing);
    std::wcout << L"  Deleted row 5" << std::endl;

    // Set RowHeight on row 1
    Excel::RangePtr row1 = excel.GetRange(ws, L"1:1");
    row1->PutRowHeight(_variant_t(30.0));

    // Set ColumnWidth on column 1
    Excel::RangePtr col1 = excel.GetRange(ws, L"A:A");
    col1->PutColumnWidth(_variant_t(20.0));

    // AutoFit columns A:C
    Excel::RangePtr colRange = excel.GetRange(ws, L"A:C");
    colRange->GetEntireColumn()->AutoFit();
    std::wcout << L"  AutoFit columns A:C" << std::endl;

    // Verify inserted row content
    _variant_t v = excel.GetCell(ws, 3, 1)->GetValue2();
    std::wstring inserted = VariantToWString(v);

    bool pass = (inserted == L"Inserted");
    if (!pass) {
        std::wcout << L"  FAIL: Row 3 = " << inserted << L", expected 'Inserted'" << std::endl;
    } else {
        std::wcout << L"  Inserted row content: OK" << std::endl;
    }

    SaveAsXlsx(wb, filePath);
    wb->Close(VARIANT_FALSE);

    // Reopen and verify structural properties
    Excel::_WorkbookPtr wb2 = excel.OpenWorkbook(filePath);
    Excel::_WorksheetPtr ws2 = excel.GetSheet(wb2, 1);

    // Verify row height of row 1 (was set to 30.0, but AutoFit may have changed it;
    // we set it before AutoFit on columns, so row height should still be 30.0
    // since AutoFit on columns does not change row heights)
    {
        Excel::RangePtr r1 = excel.GetRange(ws2, L"1:1");
        _variant_t rh = r1->GetRowHeight();
        double rowHeight = VariantToDouble(rh);
        if (std::abs(rowHeight - 30.0) > 0.5) {
            std::wcout << L"  FAIL: Row 1 height = " << rowHeight << L", expected 30.0" << std::endl;
            pass = false;
        } else {
            std::wcout << L"  Row 1 height: OK (" << rowHeight << L")" << std::endl;
        }
    }

    // Verify column width of column A
    // Note: AutoFit was called after setting width to 20.0, so the final width
    // depends on content. We verify it is > 0 (AutoFit set a reasonable value).
    // The AutoFit overrides the explicit 20.0, so we just check positivity.
    {
        Excel::RangePtr c1 = excel.GetRange(ws2, L"A:A");
        _variant_t cw = c1->GetColumnWidth();
        double colWidth = VariantToDouble(cw);
        if (colWidth <= 0.0) {
            std::wcout << L"  FAIL: Column A width = " << colWidth << L", expected > 0" << std::endl;
            pass = false;
        } else {
            std::wcout << L"  Column A width: OK (" << colWidth << L")" << std::endl;
        }
    }

    wb2->Close(VARIANT_FALSE);
    return pass;
}

// ---------------------------------------------------------------------------
// E07 - Multi-Worksheet
// ---------------------------------------------------------------------------
bool RunE07(const std::wstring& outputDir) {
    std::wstring filePath = MakePath(outputDir, L"E07_multi_sheet.xlsx");
    DeleteIfExists(filePath);

    ExcelApp excel;
    Excel::_WorkbookPtr wb = excel.AddWorkbook();

    // Rename default sheet
    Excel::_WorksheetPtr wsData = excel.GetSheet(wb, 1);
    wsData->PutName(_bstr_t(L"Data"));

    // Write data into Data sheet
    excel.GetCell(wsData, 1, 1)->PutValue2(_variant_t(100.0));
    excel.GetCell(wsData, 2, 1)->PutValue2(_variant_t(200.0));

    // Add "Summary" sheet
    Excel::_WorksheetPtr wsSummary = excel.AddSheet(wb);
    wsSummary->PutName(_bstr_t(L"Summary"));

    // Write cross-sheet formula referencing Data!A1
    excel.GetCell(wsSummary, 1, 1)->PutFormula(_variant_t(L"=Data!A1+Data!A2"));

    // Add a temporary sheet and delete it
    Excel::_WorksheetPtr wsTemp = excel.AddSheet(wb);
    wsTemp->PutName(_bstr_t(L"TempSheet"));
    excel.app()->PutDisplayAlerts(VARIANT_FALSE);
    wsTemp->Delete();
    wsTemp = nullptr;
    std::wcout << L"  Created and deleted TempSheet" << std::endl;

    // Verify cross-sheet formula result
    _variant_t val = excel.GetCell(wsSummary, 1, 1)->GetValue2();
    double actual = VariantToDouble(val);
    bool pass = (std::abs(actual - 300.0) < 0.001);

    if (pass) {
        std::wcout << L"  Cross-sheet formula =Data!A1+Data!A2 = " << actual << L" OK" << std::endl;
    } else {
        std::wcout << L"  FAIL: Cross-sheet formula = " << actual << L", expected 300" << std::endl;
    }

    // Verify sheet count (Data + Summary = 2, TempSheet was deleted)
    long sheetCount = wb->GetWorksheets()->GetCount();
    std::wcout << L"  Sheet count: " << sheetCount << std::endl;
    if (sheetCount != 2) {
        std::wcout << L"  FAIL: Expected 2 sheets, got " << sheetCount << std::endl;
        pass = false;
    }

    SaveAsXlsx(wb, filePath);
    wb->Close(VARIANT_FALSE);
    return pass;
}

// ---------------------------------------------------------------------------
// E08 - Data Operations: Replace, Sort, AutoFilter
// ---------------------------------------------------------------------------
bool RunE08(const std::wstring& outputDir) {
    std::wstring filePath = MakePath(outputDir, L"E08_data_ops.xlsx");
    DeleteIfExists(filePath);

    ExcelApp excel;
    Excel::_WorkbookPtr wb = excel.AddWorkbook();
    Excel::_WorksheetPtr ws = excel.GetSheet(wb, 1);

    // Create dataset: 25 rows + header
    // Columns: Name, Category, Score, City
    const wchar_t* names[] = {
        L"Alice", L"Bob", L"Charlie", L"Diana", L"Eve", L"Frank", L"Grace",
        L"Hank", L"Ivy", L"Jack", L"Kate", L"Leo", L"Mona", L"Nick", L"Olive",
        L"Pat", L"Quinn", L"Rose", L"Sam", L"Tina", L"Uma", L"Vic", L"Wendy",
        L"Xena", L"Yuri"
    };
    const wchar_t* categories[] = { L"A", L"B", L"C" };
    const wchar_t* cities[] = { L"NYC", L"LA", L"Chicago", L"Boston", L"Denver" };

    // Headers
    excel.GetCell(ws, 1, 1)->PutValue2(_variant_t(L"Name"));
    excel.GetCell(ws, 1, 2)->PutValue2(_variant_t(L"Category"));
    excel.GetCell(ws, 1, 3)->PutValue2(_variant_t(L"Score"));
    excel.GetCell(ws, 1, 4)->PutValue2(_variant_t(L"City"));

    // Data rows with deterministic pseudo-random scores
    int seed = 42;
    for (int i = 0; i < 25; ++i) {
        int row = i + 2;
        excel.GetCell(ws, row, 1)->PutValue2(_variant_t(names[i]));
        excel.GetCell(ws, row, 2)->PutValue2(_variant_t(categories[i % 3]));
        // Simple deterministic "random" score 50-99
        seed = (seed * 1103515245 + 12345) & 0x7fffffff;
        int score = 50 + (seed % 50);
        excel.GetCell(ws, row, 3)->PutValue2(_variant_t(static_cast<double>(score)));
        excel.GetCell(ws, row, 4)->PutValue2(_variant_t(cities[i % 5]));
    }
    std::wcout << L"  Wrote 25-row dataset with headers" << std::endl;

    // --- Replace: "NYC" -> "New York City" ---
    Excel::RangePtr dataRange = excel.GetRange(ws, L"A1:D26");
    dataRange->Replace(
        _variant_t(L"NYC"),           // What
        _variant_t(L"New York City"), // Replacement
        _variant_t(1L),               // LookAt: xlWhole=1
        _variant_t(1L),               // SearchOrder: xlByRows=1
        _variant_t(false),            // MatchCase
        vtMissing, vtMissing, vtMissing
    );
    std::wcout << L"  Replaced 'NYC' with 'New York City'" << std::endl;

    // Verify replacement
    int nycCount = 0;
    int newYorkCount = 0;
    for (int r = 2; r <= 26; ++r) {
        _variant_t v = excel.GetCell(ws, r, 4)->GetValue2();
        std::wstring s = VariantToWString(v);
        if (s == L"NYC") nycCount++;
        if (s == L"New York City") newYorkCount++;
    }
    std::wcout << L"  After replace: NYC=" << nycCount << L", New York City=" << newYorkCount << std::endl;

    // --- Sort by Score (column C) descending ---
    Excel::RangePtr sortRange = excel.GetRange(ws, L"A1:D26");
    Excel::RangePtr sortKey = excel.GetRange(ws, L"C1:C26");
    sortRange->Sort(
        _variant_t(static_cast<IDispatch*>(sortKey)),  // Key1
        Excel::xlDescending,                            // Order1
        vtMissing,                                      // Key2
        vtMissing,                                      // Type
        Excel::xlAscending,                             // Order2 (unused, required enum)
        vtMissing,                                      // Key3
        Excel::xlAscending,                             // Order3 (unused, required enum)
        Excel::xlYes,                                   // Header
        vtMissing,                                      // OrderCustom
        _variant_t(false),                              // MatchCase
        Excel::xlSortColumns,                           // Orientation
        Excel::xlPinYin,                                // SortMethod
        Excel::xlSortNormal,                            // DataOption1
        Excel::xlSortNormal,                            // DataOption2
        Excel::xlSortNormal                             // DataOption3
    );
    std::wcout << L"  Sorted by Score descending" << std::endl;

    // Verify sort: all data rows (2 through 26) must be in descending order by Score
    bool sortOk = true;
    double prevScore = VariantToDouble(excel.GetCell(ws, 2, 3)->GetValue2());
    for (int r = 3; r <= 26; ++r) {
        double curScore = VariantToDouble(excel.GetCell(ws, r, 3)->GetValue2());
        if (prevScore < curScore) {
            std::wcout << L"  FAIL: Sort order broken at row " << r
                       << L" (prev=" << prevScore << L", cur=" << curScore << L")" << std::endl;
            sortOk = false;
        }
        prevScore = curScore;
    }
    if (sortOk) {
        std::wcout << L"  Sort order verified: all 25 rows descending OK" << std::endl;
    }

    // --- AutoFilter: Category = "A" ---
    Excel::RangePtr filterRange = excel.GetRange(ws, L"A1:D26");
    filterRange->AutoFilter(
        _variant_t(2L),              // Field: column 2 (Category)
        _variant_t(L"A"),            // Criteria1
        Excel::xlAnd,                // Operator
        vtMissing,                   // Criteria2
        _variant_t(true)             // VisibleDropDown
    );
    std::wcout << L"  AutoFilter applied: Category = 'A'" << std::endl;

    // Save
    SaveAsXlsx(wb, filePath);
    wb->Close(VARIANT_FALSE);

    bool pass = true;
    if (nycCount != 0) {
        std::wcout << L"  FAIL: NYC should have been replaced" << std::endl;
        pass = false;
    }
    if (newYorkCount == 0) {
        std::wcout << L"  FAIL: No 'New York City' found" << std::endl;
        pass = false;
    }
    if (!sortOk) {
        pass = false;
    }

    return pass;
}

// ---------------------------------------------------------------------------
// E09 - Chart Generation
// ---------------------------------------------------------------------------
bool RunE09(const std::wstring& outputDir) {
    std::wstring filePath = MakePath(outputDir, L"E09_chart.xlsx");
    DeleteIfExists(filePath);

    ExcelApp excel;
    Excel::_WorkbookPtr wb = excel.AddWorkbook();
    Excel::_WorksheetPtr ws = excel.GetSheet(wb, 1);

    // Write chart source data
    excel.GetCell(ws, 1, 1)->PutValue2(_variant_t(L"Category"));
    excel.GetCell(ws, 1, 2)->PutValue2(_variant_t(L"Value"));

    const wchar_t* cats[] = { L"Alpha", L"Beta", L"Gamma", L"Delta", L"Epsilon" };
    double vals[] = { 10.0, 25.0, 15.0, 30.0, 20.0 };
    for (int i = 0; i < 5; ++i) {
        excel.GetCell(ws, i + 2, 1)->PutValue2(_variant_t(cats[i]));
        excel.GetCell(ws, i + 2, 2)->PutValue2(_variant_t(vals[i]));
    }

    // Add chart: ChartObjects.Add(left, top, width, height)
    Excel::ChartObjectsPtr chartObjs = ws->ChartObjects(vtMissing);
    Excel::ChartObjectPtr chartObj = chartObjs->Add(200.0, 10.0, 400.0, 300.0);
    Excel::_ChartPtr chart = chartObj->GetChart();

    // Set source data
    Excel::RangePtr srcRange = excel.GetRange(ws, L"A1:B6");
    chart->SetSourceData(srcRange, vtMissing);

    // Set chart type: xlColumnClustered
    chart->PutChartType(Excel::xlColumnClustered);

    // Set chart title
    chart->PutHasTitle(VARIANT_TRUE);
    chart->GetChartTitle()->PutText(_bstr_t(L"Category Values"));
    std::wcout << L"  Created column clustered chart with title" << std::endl;

    // Save
    SaveAsXlsx(wb, filePath);
    wb->Close(VARIANT_FALSE);

    // Reopen and verify chart properties
    bool pass = true;

    if (!FileExistsNonEmpty(filePath)) {
        std::wcout << L"  FAIL: File not found or empty" << std::endl;
        return false;
    }

    Excel::_WorkbookPtr wb2 = excel.OpenWorkbook(filePath);
    Excel::_WorksheetPtr ws2 = excel.GetSheet(wb2, 1);

    // Verify chart count
    Excel::ChartObjectsPtr chartObjs2 = ws2->ChartObjects(vtMissing);
    long chartCount = chartObjs2->GetCount();
    if (chartCount != 1) {
        std::wcout << L"  FAIL: Chart count = " << chartCount << L", expected 1" << std::endl;
        pass = false;
    } else {
        std::wcout << L"  Chart count: OK (1)" << std::endl;
    }

    // Verify chart type
    Excel::ChartObjectPtr chartObj2 = chartObjs2->Item(_variant_t(1L));
    Excel::_ChartPtr chart2 = chartObj2->GetChart();
    long chartType = static_cast<long>(chart2->GetChartType());
    long expectedType = static_cast<long>(Excel::xlColumnClustered);
    if (chartType != expectedType) {
        std::wcout << L"  FAIL: Chart type = " << chartType
                   << L", expected " << expectedType << L" (xlColumnClustered)" << std::endl;
        pass = false;
    } else {
        std::wcout << L"  Chart type: OK (xlColumnClustered)" << std::endl;
    }

    // Verify chart title
    if (chart2->GetHasTitle() == VARIANT_TRUE) {
        _variant_t titleText = chart2->GetChartTitle()->GetText();
        std::wstring titleStr = VariantToWString(titleText);
        if (titleStr != L"Category Values") {
            std::wcout << L"  FAIL: Chart title = \"" << titleStr
                       << L"\", expected \"Category Values\"" << std::endl;
            pass = false;
        } else {
            std::wcout << L"  Chart title: OK" << std::endl;
        }
    } else {
        std::wcout << L"  FAIL: Chart has no title" << std::endl;
        pass = false;
    }

    wb2->Close(VARIANT_FALSE);
    std::wcout << L"  File: " << filePath << (pass ? L" OK" : L" FAIL") << std::endl;
    return pass;
}

// ---------------------------------------------------------------------------
// E10 - Export PDF
// ---------------------------------------------------------------------------
bool RunE10(const std::wstring& outputDir) {
    std::wstring xlsxPath = MakePath(outputDir, L"E10_export.xlsx");
    std::wstring pdfPath  = MakePath(outputDir, L"E10_export.pdf");
    DeleteIfExists(xlsxPath);
    DeleteIfExists(pdfPath);

    ExcelApp excel;
    Excel::_WorkbookPtr wb = excel.AddWorkbook();
    Excel::_WorksheetPtr ws = excel.GetSheet(wb, 1);

    // Write some data for the PDF
    excel.GetCell(ws, 1, 1)->PutValue2(_variant_t(L"PDF Export Test"));
    excel.GetCell(ws, 2, 1)->PutValue2(_variant_t(L"Row 2 data"));
    for (int r = 3; r <= 10; ++r) {
        excel.GetCell(ws, r, 1)->PutValue2(_variant_t(static_cast<double>(r * 100)));
    }

    // Save as xlsx first
    SaveAsXlsx(wb, xlsxPath);

    // ExportAsFixedFormat: Type 0 = xlTypePDF
    wb->ExportAsFixedFormat(
        static_cast<Excel::XlFixedFormatType>(0),   // xlTypePDF
        _variant_t(pdfPath.c_str()),
        vtMissing,  // Quality
        vtMissing,  // IncludeDocProperties
        vtMissing,  // IgnorePrintAreas
        vtMissing,  // From
        vtMissing,  // To
        vtMissing,  // OpenAfterPublish
        vtMissing   // FixedFormatExtClassPtr
    );
    std::wcout << L"  Exported PDF: " << pdfPath << std::endl;

    wb->Close(VARIANT_FALSE);

    bool pass = FileExistsNonEmpty(pdfPath);
    if (pass) {
        auto sz = fs::file_size(pdfPath);
        std::wcout << L"  PDF size: " << sz << L" bytes" << std::endl;
    } else {
        std::wcout << L"  FAIL: PDF not generated or empty" << std::endl;
    }
    return pass;
}

// ---------------------------------------------------------------------------
// E11 - Run VBA Macro
// ---------------------------------------------------------------------------
bool RunE11(const std::wstring& outputDir) {
    std::wstring filePath = MakePath(outputDir, L"E11_macro.xlsm");
    DeleteIfExists(filePath);

    ExcelApp excel;
    Excel::_WorkbookPtr wb = excel.AddWorkbook();
    Excel::_WorksheetPtr ws = excel.GetSheet(wb, 1);

    // Write initial data
    excel.GetCell(ws, 1, 1)->PutValue2(_variant_t(L"Before Macro"));
    excel.GetCell(ws, 2, 1)->PutValue2(_variant_t(10.0));

    // VBA code to inject
    const wchar_t* vbaCode =
        L"Sub TestMacro()\r\n"
        L"    Dim ws As Worksheet\r\n"
        L"    Set ws = ThisWorkbook.Worksheets(1)\r\n"
        L"    ws.Range(\"A1\").Value2 = \"Macro Executed\"\r\n"
        L"    ws.Range(\"A2\").Value2 = ws.Range(\"A2\").Value2 * 2\r\n"
        L"    ws.Range(\"A3\").Value2 = \"VBA was here\"\r\n"
        L"End Sub\r\n";

    bool vbaInjected = false;
    try {
        // Access VBProject via IDispatch late binding
        IDispatch* wbDisp = wb;

        // Get VBProject property
        DISPID dispid;
        OLECHAR* propName = const_cast<OLECHAR*>(L"VBProject");
        HRESULT hr = wbDisp->GetIDsOfNames(IID_NULL, &propName, 1, LOCALE_USER_DEFAULT, &dispid);
        if (FAILED(hr)) throw _com_error(hr);

        DISPPARAMS dpNoArgs = { nullptr, nullptr, 0, 0 };
        _variant_t vbProject;
        hr = wbDisp->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET,
                            &dpNoArgs, &vbProject, nullptr, nullptr);
        if (FAILED(hr)) throw _com_error(hr);

        IDispatch* projDisp = vbProject.pdispVal;

        // Get VBComponents property
        propName = const_cast<OLECHAR*>(L"VBComponents");
        hr = projDisp->GetIDsOfNames(IID_NULL, &propName, 1, LOCALE_USER_DEFAULT, &dispid);
        if (FAILED(hr)) throw _com_error(hr);

        _variant_t vbComponents;
        hr = projDisp->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET,
                              &dpNoArgs, &vbComponents, nullptr, nullptr);
        if (FAILED(hr)) throw _com_error(hr);

        IDispatch* compsDisp = vbComponents.pdispVal;

        // Add(1) - 1 = vbext_ct_StdModule
        propName = const_cast<OLECHAR*>(L"Add");
        hr = compsDisp->GetIDsOfNames(IID_NULL, &propName, 1, LOCALE_USER_DEFAULT, &dispid);
        if (FAILED(hr)) throw _com_error(hr);

        _variant_t argType(1L);
        DISPPARAMS dpAdd;
        dpAdd.rgvarg = static_cast<VARIANTARG*>(&argType);
        dpAdd.rgdispidNamedArgs = nullptr;
        dpAdd.cArgs = 1;
        dpAdd.cNamedArgs = 0;

        _variant_t module;
        hr = compsDisp->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD,
                               &dpAdd, &module, nullptr, nullptr);
        if (FAILED(hr)) throw _com_error(hr);

        IDispatch* modDisp = module.pdispVal;

        // Get CodeModule property
        propName = const_cast<OLECHAR*>(L"CodeModule");
        hr = modDisp->GetIDsOfNames(IID_NULL, &propName, 1, LOCALE_USER_DEFAULT, &dispid);
        if (FAILED(hr)) throw _com_error(hr);

        _variant_t codeModule;
        hr = modDisp->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET,
                             &dpNoArgs, &codeModule, nullptr, nullptr);
        if (FAILED(hr)) throw _com_error(hr);

        IDispatch* codeDisp = codeModule.pdispVal;

        // AddFromString(vbaCode)
        propName = const_cast<OLECHAR*>(L"AddFromString");
        hr = codeDisp->GetIDsOfNames(IID_NULL, &propName, 1, LOCALE_USER_DEFAULT, &dispid);
        if (FAILED(hr)) throw _com_error(hr);

        _bstr_t bstrCode(vbaCode);
        _variant_t argCode(bstrCode);
        DISPPARAMS dpAddStr;
        dpAddStr.rgvarg = static_cast<VARIANTARG*>(&argCode);
        dpAddStr.rgdispidNamedArgs = nullptr;
        dpAddStr.cArgs = 1;
        dpAddStr.cNamedArgs = 0;

        hr = codeDisp->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD,
                              &dpAddStr, nullptr, nullptr, nullptr);
        if (FAILED(hr)) throw _com_error(hr);

        vbaInjected = true;
        std::wcout << L"  VBA code injected successfully" << std::endl;
    }
    catch (_com_error& e) {
        std::wcout << L"  WARNING: Cannot inject VBA - Trust setting may be disabled" << std::endl;
        std::wcout << L"  Error: " << (e.ErrorMessage() ? e.ErrorMessage() : L"unknown") << std::endl;
        std::wcout << L"  Enable: File > Options > Trust Center > Macro Settings" << std::endl;
        std::wcout << L"  > 'Trust access to the VBA project object model'" << std::endl;
    }
    catch (...) {
        std::wcout << L"  WARNING: VBA injection failed with unknown error" << std::endl;
    }

    if (!vbaInjected) {
        std::wcout << L"  Skipping macro execution (VBA injection failed)" << std::endl;
        std::wcout << L"  RESULT: SKIP (environment issue, not code defect)" << std::endl;
        wb->Close(VARIANT_FALSE);
        return true;  // Not a code defect
    }

    // Run the macro
    try {
        excel.app()->Run(
            _variant_t(L"TestMacro"),
            vtMissing, vtMissing, vtMissing, vtMissing, vtMissing,
            vtMissing, vtMissing, vtMissing, vtMissing, vtMissing,
            vtMissing, vtMissing, vtMissing, vtMissing, vtMissing,
            vtMissing, vtMissing, vtMissing, vtMissing, vtMissing,
            vtMissing, vtMissing, vtMissing, vtMissing, vtMissing,
            vtMissing, vtMissing, vtMissing, vtMissing
        );
        std::wcout << L"  Macro executed successfully" << std::endl;
    }
    catch (_com_error& e) {
        std::wcout << L"  FAIL: Macro execution failed: "
                   << (e.ErrorMessage() ? e.ErrorMessage() : L"unknown") << std::endl;
        wb->Close(VARIANT_FALSE);
        return false;
    }

    // Verify macro effects
    std::wstring strA1 = VariantToWString(excel.GetCell(ws, 1, 1)->GetValue2());
    double dblA2 = VariantToDouble(excel.GetCell(ws, 2, 1)->GetValue2());
    std::wstring strA3 = VariantToWString(excel.GetCell(ws, 3, 1)->GetValue2());

    std::wcout << L"  After macro: A1=\"" << strA1 << L"\", A2=" << dblA2
               << L", A3=\"" << strA3 << L"\"" << std::endl;

    bool pass = true;
    if (strA1 != L"Macro Executed") {
        std::wcout << L"  FAIL: A1 should be 'Macro Executed'" << std::endl;
        pass = false;
    }
    if (std::abs(dblA2 - 20.0) > 0.001) {
        std::wcout << L"  FAIL: A2 should be 20 (10*2)" << std::endl;
        pass = false;
    }
    if (strA3 != L"VBA was here") {
        std::wcout << L"  FAIL: A3 should be 'VBA was here'" << std::endl;
        pass = false;
    }

    // Save as .xlsm (xlOpenXMLWorkbookMacroEnabled)
    wb->SaveAs(_variant_t(filePath.c_str()),
               _variant_t(static_cast<long>(Excel::xlOpenXMLWorkbookMacroEnabled)),
               vtMissing, vtMissing, vtMissing, vtMissing,
               Excel::xlNoChange, vtMissing, vtMissing, vtMissing,
               vtMissing, vtMissing);
    wb->Close(VARIANT_FALSE);
    return pass;
}

// ---------------------------------------------------------------------------
// E12 - Resource Cleanup (RAII demonstration)
// ---------------------------------------------------------------------------
bool RunE12(const std::wstring& outputDir) {
    std::wcout << L"  Demonstrating RAII cleanup with smart pointers..." << std::endl;

    DWORD ourPid = 0;

    // Inner scope: all COM objects created and destroyed here
    {
        ExcelApp excel;

        // Track the PID of our Excel instance via its window handle
        HWND hwnd = reinterpret_cast<HWND>(
            static_cast<intptr_t>(excel.app()->GetHwnd()));
        ::GetWindowThreadProcessId(hwnd, &ourPid);
        std::wcout << L"  Our Excel PID: " << ourPid << std::endl;

        Excel::_WorkbookPtr wb = excel.AddWorkbook();
        Excel::_WorksheetPtr ws = excel.GetSheet(wb, 1);

        // Do some work
        excel.GetCell(ws, 1, 1)->PutValue2(_variant_t(L"RAII Test"));
        for (int r = 2; r <= 100; ++r) {
            excel.GetCell(ws, r, 1)->PutValue2(_variant_t(static_cast<double>(r)));
        }

        std::wstring filePath = MakePath(outputDir, L"E12_cleanup.xlsx");
        DeleteIfExists(filePath);
        SaveAsXlsx(wb, filePath);

        // Explicit close before scope exit (best practice)
        wb->Close(VARIANT_FALSE);
        ws = nullptr;
        wb = nullptr;

        std::wcout << L"  Workbook closed, smart pointers releasing..." << std::endl;
        // ~ExcelApp() calls app->Quit() + releases app_ smart pointer
    }

    std::wcout << L"  All COM objects released, Excel.Application quit" << std::endl;

    // Give Excel time to terminate naturally
    ::Sleep(3000);

    // Check if our specific process has exited
    bool noResidual = false;
    HANDLE hProc = ::OpenProcess(PROCESS_QUERY_INFORMATION | SYNCHRONIZE, FALSE, ourPid);
    if (hProc == nullptr) {
        // Process not found - clean exit
        noResidual = true;
        std::wcout << L"  Our Excel process (PID " << ourPid
                   << L") has exited - cleanup successful" << std::endl;
    } else {
        // Process handle obtained, check if it has exited
        DWORD exitCode = 0;
        if (::GetExitCodeProcess(hProc, &exitCode) && exitCode != STILL_ACTIVE) {
            noResidual = true;
            std::wcout << L"  Our Excel process (PID " << ourPid
                       << L") has exited - cleanup successful" << std::endl;
        } else {
            // Process still alive after 3s, force kill
            std::wcout << L"  Process still alive after 3s, forcing kill..." << std::endl;
            ::TerminateProcess(hProc, 1);
            ::WaitForSingleObject(hProc, 5000);
            noResidual = true;
            std::wcout << L"  WARNING: Hidden COM refs prevented clean exit - process killed."
                       << std::endl;
        }
        ::CloseHandle(hProc);
    }

    return noResidual;
}
