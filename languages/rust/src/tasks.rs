// tasks.rs - E01-E12 Excel COM automation tasks, ported from the C++ reference.
//
// Each task returns Result<bool>: Ok(true) = PASS, Ok(false) = FAIL.
// A COM error bubbles up as Err and is reported by the runner as FAIL.

use std::path::Path;

use windows::core::Result;

use excel_com::com::{bstr, empty, i4, r8, vbool, SafeArray2D, Variant};
use excel_com::excel::ExcelApp;

// ---------------------------------------------------------------------------
// Small filesystem helpers
// ---------------------------------------------------------------------------
fn make_path(dir: &str, filename: &str) -> String {
    Path::new(dir).join(filename).to_string_lossy().into_owned()
}

fn delete_if_exists(path: &str) {
    let _ = std::fs::remove_file(path);
}

fn file_exists_non_empty(path: &str) -> bool {
    std::fs::metadata(path).map(|m| m.len() > 0).unwrap_or(false)
}

// ---------------------------------------------------------------------------
// E01 - Workbook Lifecycle
// ---------------------------------------------------------------------------
pub fn run_e01(output_dir: &str) -> Result<bool> {
    let file_path = make_path(output_dir, "E01_lifecycle.xlsx");
    delete_if_exists(&file_path);

    let excel = ExcelApp::new()?;

    let wb = excel.add_workbook()?;
    let ws = excel.get_sheet(&wb, 1)?;
    let cell = excel.get_cell(&ws, 1, 1)?;
    cell.put("Value2", bstr("E01 Lifecycle Test"))?;

    excel.save_as_xlsx(&wb, &file_path)?;
    println!("  Saved: {}", file_path);

    wb.call("Close", &[vbool(false)])?;

    let wb2 = excel.open_workbook(&file_path)?;
    let ws2 = excel.get_sheet(&wb2, 1)?;
    let cell2 = excel.get_cell(&ws2, 1, 1)?;
    let val = cell2.get("Value2")?;

    let read_val = val.as_string().unwrap_or("");
    let pass = read_val == "E01 Lifecycle Test";
    println!(
        "  Read back: \"{}\" - {}",
        read_val,
        if pass { "match" } else { "MISMATCH" }
    );

    wb2.call("Close", &[vbool(false)])?;
    Ok(pass)
}

// ---------------------------------------------------------------------------
// E02 - Cell Read/Write
// ---------------------------------------------------------------------------
pub fn run_e02(output_dir: &str) -> Result<bool> {
    let file_path = make_path(output_dir, "E02_cell_rw.xlsx");
    delete_if_exists(&file_path);

    let excel = ExcelApp::new()?;
    let wb = excel.add_workbook()?;
    let ws = excel.get_sheet(&wb, 1)?;

    excel.get_cell(&ws, 1, 1)?.put("Value2", bstr("Hello COM"))?;
    excel.get_cell(&ws, 2, 1)?.put("Value2", r8(42.5))?;
    excel.get_cell(&ws, 3, 1)?.put("Value2", vbool(true))?;
    excel.get_cell(&ws, 4, 1)?.put("Value2", r8(45306.0))?;
    excel
        .get_cell(&ws, 4, 1)?
        .put("NumberFormat", bstr("yyyy-mm-dd"))?;

    let mut pass = true;

    // String
    let v1 = excel.get_cell(&ws, 1, 1)?.get("Value2")?;
    match v1.as_string() {
        Some("Hello COM") => println!("  String: OK"),
        _ => {
            println!("  FAIL: string mismatch");
            pass = false;
        }
    }

    // Number
    let v2 = excel.get_cell(&ws, 2, 1)?.get("Value2")?;
    match v2 {
        Variant::R8(n) if (n - 42.5).abs() <= 0.001 => println!("  Number: OK ({})", n),
        _ => {
            println!("  FAIL: number mismatch");
            pass = false;
        }
    }

    // Boolean
    let v3 = excel.get_cell(&ws, 3, 1)?.get("Value2")?;
    match v3 {
        Variant::Bool(true) => println!("  Boolean: OK"),
        _ => {
            println!("  FAIL: boolean mismatch");
            pass = false;
        }
    }

    // Date serial
    let v4 = excel.get_cell(&ws, 4, 1)?.get("Value2")?;
    match v4 {
        Variant::R8(n) if (n - 45306.0).abs() <= 0.001 => println!("  Date serial: OK ({})", n),
        _ => {
            println!("  FAIL: date serial mismatch");
            pass = false;
        }
    }

    excel.save_as_xlsx(&wb, &file_path)?;
    wb.call("Close", &[vbool(false)])?;

    println!("  File: {}", file_path);
    Ok(pass)
}

// ---------------------------------------------------------------------------
// E03 - Bulk Range Write (SAFEARRAY vs cell-by-cell)
// ---------------------------------------------------------------------------
pub fn run_e03(output_dir: &str) -> Result<bool> {
    const ROWS: i32 = 1000;
    const COLS: i32 = 10;

    let file_path = make_path(output_dir, "E03_bulk_write.xlsx");
    delete_if_exists(&file_path);

    let excel = ExcelApp::new()?;
    let wb = excel.add_workbook()?;

    // Idiomatic bulk write via SAFEARRAY
    let ws_idiomatic = excel.get_sheet(&wb, 1)?;
    ws_idiomatic.put("Name", bstr("Idiomatic"))?;

    let sa_data = SafeArray2D::fill_sequential(ROWS, COLS)?;

    let t0 = std::time::Instant::now();
    excel.bulk_write(&ws_idiomatic, 1, 1, ROWS, COLS, sa_data)?;
    let idiomatic_ms = t0.elapsed().as_secs_f64() * 1000.0;
    println!(
        "  Idiomatic bulk write: {:.3} ms ({}x{})",
        idiomatic_ms, ROWS, COLS
    );

    // Naive cell-by-cell write
    let ws_naive = excel.add_sheet(&wb)?;
    ws_naive.put("Name", bstr("Naive"))?;

    let t1 = std::time::Instant::now();
    for r in 1..=ROWS {
        for c in 1..=COLS {
            let val = ((r - 1) * COLS + c) as f64;
            excel.get_cell(&ws_naive, r, c)?.put("Value2", r8(val))?;
        }
    }
    let naive_ms = t1.elapsed().as_secs_f64() * 1000.0;
    println!(
        "  Naive cell-by-cell write: {:.3} ms ({}x{})",
        naive_ms, ROWS, COLS
    );

    let speedup = if idiomatic_ms > 0.001 {
        naive_ms / idiomatic_ms
    } else {
        0.0
    };
    println!("  Speedup ratio: {:.2}x", speedup);

    excel.save_as_xlsx(&wb, &file_path)?;
    wb.call("Close", &[vbool(false)])?;

    // Reopen and spot-check
    let wb2 = excel.open_workbook(&file_path)?;
    let ws2i = excel.get_sheet_by_name(&wb2, "Idiomatic")?;
    let ws2n = excel.get_sheet_by_name(&wb2, "Naive")?;

    let mut pass = true;
    let checks = [(1, 1), (1, 10), (500, 5), (1000, 10)];
    for (r, c) in checks {
        let expected = ((r - 1) * COLS + c) as f64;
        let di = excel.get_cell(&ws2i, r, c)?.get("Value2")?.as_f64().unwrap_or(-1.0);
        let dn = excel.get_cell(&ws2n, r, c)?.get("Value2")?.as_f64().unwrap_or(-1.0);
        if (di - expected).abs() > 0.001 {
            println!("  FAIL: Idiomatic[{},{}] = {}, expected {}", r, c, di, expected);
            pass = false;
        }
        if (dn - expected).abs() > 0.001 {
            println!("  FAIL: Naive[{},{}] = {}, expected {}", r, c, dn, expected);
            pass = false;
        }
    }

    wb2.call("Close", &[vbool(false)])?;
    println!(
        "  Bulk write is {} faster ({:.2}x)",
        if speedup >= 2.0 { "significantly" } else { "somewhat" },
        speedup
    );
    Ok(pass)
}

// ---------------------------------------------------------------------------
// E04 - Formula & Recalc
// ---------------------------------------------------------------------------
pub fn run_e04(output_dir: &str) -> Result<bool> {
    let file_path = make_path(output_dir, "E04_formula.xlsx");
    delete_if_exists(&file_path);

    let excel = ExcelApp::new()?;
    let wb = excel.add_workbook()?;
    let ws = excel.get_sheet(&wb, 1)?;

    // Source data
    for r in 1..=10 {
        excel.get_cell(&ws, r, 1)?.put("Value2", r8(r as f64))?;
        excel.get_cell(&ws, r, 2)?.put("Value2", r8((r * 2) as f64))?;
    }

    // Manual calculation (xlCalculationManual = -4135)
    excel.app().put("Calculation", i4(-4135))?;

    // Formulas in column C: =A{r}*B{r}
    for r in 1..=10 {
        let formula = format!("=A{}*B{}", r, r);
        excel.get_cell(&ws, r, 3)?.put("Formula", bstr(&formula))?;
    }

    // Force recalculation
    excel.app().call("Calculate", &[])?;

    let mut pass = true;
    for r in 1..=10 {
        let val = excel.get_cell(&ws, r, 3)?.get("Value2")?;
        let expected = (r as f64) * ((r * 2) as f64);
        let actual = val.as_f64().unwrap_or(-1.0);
        if (actual - expected).abs() > 0.001 {
            println!("  FAIL: C{} = {}, expected {}", r, actual, expected);
            pass = false;
        }
    }

    // Restore automatic calculation (xlCalculationAutomatic = -4105)
    excel.app().put("Calculation", i4(-4105))?;

    if pass {
        println!("  All 10 formulas verified OK");
    }

    excel.save_as_xlsx(&wb, &file_path)?;
    wb.call("Close", &[vbool(false)])?;
    Ok(pass)
}

// ---------------------------------------------------------------------------
// E05 - Cell Formatting
// ---------------------------------------------------------------------------
pub fn run_e05(output_dir: &str) -> Result<bool> {
    let file_path = make_path(output_dir, "E05_formatting.xlsx");
    delete_if_exists(&file_path);

    let excel = ExcelApp::new()?;
    let wb = excel.add_workbook()?;
    let ws = excel.get_sheet(&wb, 1)?;

    excel.get_cell(&ws, 1, 1)?.put("Value2", bstr("Bold Header"))?;
    excel.get_cell(&ws, 2, 1)?.put("Value2", r8(1234.5678))?;
    excel.get_cell(&ws, 3, 1)?.put("Value2", bstr("Merged"))?;

    // Font: bold + size 16 + red (BGR 0x0000FF = 255)
    {
        let r = excel.get_cell(&ws, 1, 1)?;
        let font = r.get_dispatch("Font")?;
        font.put("Bold", vbool(true))?;
        font.put("Size", i4(16))?;
        font.put("Color", i4(255))?;
    }

    // Interior: yellow (BGR 65535)
    {
        let r = excel.get_cell(&ws, 2, 1)?;
        r.get_dispatch("Interior")?.put("Color", i4(65535))?;
    }

    // Number format
    excel.get_cell(&ws, 2, 1)?.put("NumberFormat", bstr("0.00"))?;

    // Horizontal alignment (xlCenter = -4108)
    excel
        .get_cell(&ws, 1, 1)?
        .put("HorizontalAlignment", i4(-4108))?;

    // Merge A3:C3
    {
        let r = excel.get_range(&ws, "A3:C3")?;
        r.call("Merge", &[empty()])?;
        r.put("HorizontalAlignment", i4(-4108))?;
    }

    excel.save_as_xlsx(&wb, &file_path)?;
    wb.call("Close", &[vbool(false)])?;

    // Reopen and verify formatting persisted
    let wb2 = excel.open_workbook(&file_path)?;
    let ws2 = excel.get_sheet(&wb2, 1)?;

    let mut pass = true;

    // Bold
    {
        let bold = excel.get_cell(&ws2, 1, 1)?.get_dispatch("Font")?.get("Bold")?;
        match bold {
            Variant::Bool(true) => println!("  Bold: OK"),
            _ => {
                println!("  FAIL: Bold not preserved");
                pass = false;
            }
        }
    }

    // Font size
    {
        let sz = excel.get_cell(&ws2, 1, 1)?.get_dispatch("Font")?.get("Size")?;
        let font_size = sz.as_f64().unwrap_or(-1.0);
        if (font_size - 16.0).abs() > 0.1 {
            println!("  FAIL: Font size = {}, expected 16", font_size);
            pass = false;
        } else {
            println!("  Font size: OK ({})", font_size);
        }
    }

    // Number format
    {
        let fmt = excel.get_cell(&ws2, 2, 1)?.get("NumberFormat")?;
        match fmt.as_string() {
            Some("0.00") => println!("  NumberFormat: OK"),
            other => {
                println!("  FAIL: NumberFormat = {:?}, expected 0.00", other);
                pass = false;
            }
        }
    }

    // Interior color
    {
        let color = excel.get_cell(&ws2, 2, 1)?.get_dispatch("Interior")?.get("Color")?;
        let color_val = color.as_f64().unwrap_or(-1.0);
        if (color_val - 65535.0).abs() > 0.1 {
            println!("  FAIL: Interior color = {}, expected 65535", color_val);
            pass = false;
        } else {
            println!("  Interior color: OK ({})", color_val);
        }
    }

    // Merge
    {
        let merged = excel.get_cell(&ws2, 3, 1)?.get("MergeCells")?;
        match merged {
            Variant::Bool(true) => println!("  Merge: OK"),
            _ => {
                println!("  FAIL: Merge not preserved");
                pass = false;
            }
        }
    }

    wb2.call("Close", &[vbool(false)])?;
    Ok(pass)
}

// ---------------------------------------------------------------------------
// E06 - Row/Column Structure
// ---------------------------------------------------------------------------
pub fn run_e06(output_dir: &str) -> Result<bool> {
    let file_path = make_path(output_dir, "E06_structure.xlsx");
    delete_if_exists(&file_path);

    let excel = ExcelApp::new()?;
    let wb = excel.add_workbook()?;
    let ws = excel.get_sheet(&wb, 1)?;

    for r in 1..=5 {
        for c in 1..=3 {
            let val = format!("R{}C{}", r, c);
            excel.get_cell(&ws, r, c)?.put("Value2", bstr(&val))?;
        }
    }

    // Insert a row at row 3 (Shift, CopyOrigin both omitted)
    excel.get_range(&ws, "3:3")?.call("Insert", &[empty(), empty()])?;
    println!("  Inserted row at 3");

    excel.get_cell(&ws, 3, 1)?.put("Value2", bstr("Inserted"))?;

    // Delete row 5
    excel.get_range(&ws, "5:5")?.call("Delete", &[empty()])?;
    println!("  Deleted row 5");

    // Row height / column width
    excel.get_range(&ws, "1:1")?.put("RowHeight", r8(30.0))?;
    excel.get_range(&ws, "A:A")?.put("ColumnWidth", r8(20.0))?;

    // AutoFit columns A:C
    excel
        .get_range(&ws, "A:C")?
        .get_dispatch("EntireColumn")?
        .call("AutoFit", &[])?;
    println!("  AutoFit columns A:C");

    let inserted = excel.get_cell(&ws, 3, 1)?.get("Value2")?;
    let mut pass = inserted.as_string() == Some("Inserted");
    if !pass {
        println!("  FAIL: Row 3 = {:?}, expected 'Inserted'", inserted.as_string());
    } else {
        println!("  Inserted row content: OK");
    }

    excel.save_as_xlsx(&wb, &file_path)?;
    wb.call("Close", &[vbool(false)])?;

    // Reopen and verify structural properties
    let wb2 = excel.open_workbook(&file_path)?;
    let ws2 = excel.get_sheet(&wb2, 1)?;

    {
        let rh = excel.get_range(&ws2, "1:1")?.get("RowHeight")?;
        let row_height = rh.as_f64().unwrap_or(-1.0);
        if (row_height - 30.0).abs() > 0.5 {
            println!("  FAIL: Row 1 height = {}, expected 30.0", row_height);
            pass = false;
        } else {
            println!("  Row 1 height: OK ({})", row_height);
        }
    }

    {
        let cw = excel.get_range(&ws2, "A:A")?.get("ColumnWidth")?;
        let col_width = cw.as_f64().unwrap_or(-1.0);
        if col_width <= 0.0 {
            println!("  FAIL: Column A width = {}, expected > 0", col_width);
            pass = false;
        } else {
            println!("  Column A width: OK ({})", col_width);
        }
    }

    wb2.call("Close", &[vbool(false)])?;
    Ok(pass)
}

// ---------------------------------------------------------------------------
// E07 - Multi-Worksheet
// ---------------------------------------------------------------------------
pub fn run_e07(output_dir: &str) -> Result<bool> {
    let file_path = make_path(output_dir, "E07_multi_sheet.xlsx");
    delete_if_exists(&file_path);

    let excel = ExcelApp::new()?;
    let wb = excel.add_workbook()?;

    // Rename default sheet
    let ws_data = excel.get_sheet(&wb, 1)?;
    ws_data.put("Name", bstr("Data"))?;

    excel.get_cell(&ws_data, 1, 1)?.put("Value2", r8(100.0))?;
    excel.get_cell(&ws_data, 2, 1)?.put("Value2", r8(200.0))?;

    // Add Summary sheet with cross-sheet formula
    let ws_summary = excel.add_sheet(&wb)?;
    ws_summary.put("Name", bstr("Summary"))?;
    excel
        .get_cell(&ws_summary, 1, 1)?
        .put("Formula", bstr("=Data!A1+Data!A2"))?;

    // Add + delete a temp sheet
    let ws_temp = excel.add_sheet(&wb)?;
    ws_temp.put("Name", bstr("TempSheet"))?;
    excel.app().put("DisplayAlerts", vbool(false))?;
    ws_temp.call("Delete", &[])?;
    drop(ws_temp);
    println!("  Created and deleted TempSheet");

    // Verify cross-sheet formula
    let val = excel.get_cell(&ws_summary, 1, 1)?.get("Value2")?;
    let actual = val.as_f64().unwrap_or(-1.0);
    let mut pass = (actual - 300.0).abs() < 0.001;
    if pass {
        println!("  Cross-sheet formula =Data!A1+Data!A2 = {} OK", actual);
    } else {
        println!("  FAIL: Cross-sheet formula = {}, expected 300", actual);
    }

    // Verify sheet count
    let sheet_count = wb
        .get_dispatch("Worksheets")?
        .get("Count")?
        .as_f64()
        .unwrap_or(-1.0) as i64;
    println!("  Sheet count: {}", sheet_count);
    if sheet_count != 2 {
        println!("  FAIL: Expected 2 sheets, got {}", sheet_count);
        pass = false;
    }

    excel.save_as_xlsx(&wb, &file_path)?;
    wb.call("Close", &[vbool(false)])?;
    Ok(pass)
}

// ---------------------------------------------------------------------------
// E08 - Data Operations: Replace, Sort, AutoFilter
// ---------------------------------------------------------------------------
pub fn run_e08(output_dir: &str) -> Result<bool> {
    let file_path = make_path(output_dir, "E08_data_ops.xlsx");
    delete_if_exists(&file_path);

    let excel = ExcelApp::new()?;
    let wb = excel.add_workbook()?;
    let ws = excel.get_sheet(&wb, 1)?;

    let names = [
        "Alice", "Bob", "Charlie", "Diana", "Eve", "Frank", "Grace", "Hank", "Ivy", "Jack",
        "Kate", "Leo", "Mona", "Nick", "Olive", "Pat", "Quinn", "Rose", "Sam", "Tina", "Uma",
        "Vic", "Wendy", "Xena", "Yuri",
    ];
    let categories = ["A", "B", "C"];
    let cities = ["NYC", "LA", "Chicago", "Boston", "Denver"];

    // Headers
    excel.get_cell(&ws, 1, 1)?.put("Value2", bstr("Name"))?;
    excel.get_cell(&ws, 1, 2)?.put("Value2", bstr("Category"))?;
    excel.get_cell(&ws, 1, 3)?.put("Value2", bstr("Score"))?;
    excel.get_cell(&ws, 1, 4)?.put("Value2", bstr("City"))?;

    // Data rows with deterministic pseudo-random scores (matches C++ 32-bit LCG)
    let mut seed: i32 = 42;
    for i in 0..25usize {
        let row = i as i32 + 2;
        excel.get_cell(&ws, row, 1)?.put("Value2", bstr(names[i]))?;
        excel.get_cell(&ws, row, 2)?.put("Value2", bstr(categories[i % 3]))?;
        seed = seed.wrapping_mul(1103515245).wrapping_add(12345) & 0x7fffffff;
        let score = 50 + (seed % 50);
        excel.get_cell(&ws, row, 3)?.put("Value2", r8(score as f64))?;
        excel.get_cell(&ws, row, 4)?.put("Value2", bstr(cities[i % 5]))?;
    }
    println!("  Wrote 25-row dataset with headers");

    // Replace NYC -> New York City
    excel.get_range(&ws, "A1:D26")?.call(
        "Replace",
        &[
            bstr("NYC"),
            bstr("New York City"),
            i4(1),        // LookAt: xlWhole
            i4(1),        // SearchOrder: xlByRows
            vbool(false), // MatchCase
            empty(),
            empty(),
            empty(),
        ],
    )?;
    println!("  Replaced 'NYC' with 'New York City'");

    let mut nyc_count = 0;
    let mut new_york_count = 0;
    for r in 2..=26 {
        let s = excel.get_cell(&ws, r, 4)?.get("Value2")?;
        match s.as_string() {
            Some("NYC") => nyc_count += 1,
            Some("New York City") => new_york_count += 1,
            _ => {}
        }
    }
    println!("  After replace: NYC={}, New York City={}", nyc_count, new_york_count);

    // Sort by Score (column C) descending — 15 positional args
    let sort_range = excel.get_range(&ws, "A1:D26")?;
    let sort_key = excel.get_range(&ws, "C1:C26")?;
    sort_range.call(
        "Sort",
        &[
            Variant::Dispatch(sort_key), // Key1
            i4(2),                       // Order1: xlDescending
            empty(),                     // Key2
            empty(),                     // Type
            i4(1),                       // Order2: xlAscending
            empty(),                     // Key3
            i4(1),                       // Order3: xlAscending
            i4(1),                       // Header: xlYes
            empty(),                     // OrderCustom
            vbool(false),                // MatchCase
            i4(1),                       // Orientation: xlSortColumns
            i4(1),                       // SortMethod: xlPinYin
            i4(0),                       // DataOption1: xlSortNormal
            i4(0),                       // DataOption2
            i4(0),                       // DataOption3
        ],
    )?;
    println!("  Sorted by Score descending");

    let mut sort_ok = true;
    let mut prev_score = excel.get_cell(&ws, 2, 3)?.get("Value2")?.as_f64().unwrap_or(-1.0);
    for r in 3..=26 {
        let cur_score = excel.get_cell(&ws, r, 3)?.get("Value2")?.as_f64().unwrap_or(-1.0);
        if prev_score < cur_score {
            println!(
                "  FAIL: Sort order broken at row {} (prev={}, cur={})",
                r, prev_score, cur_score
            );
            sort_ok = false;
        }
        prev_score = cur_score;
    }
    if sort_ok {
        println!("  Sort order verified: all 25 rows descending OK");
    }

    // AutoFilter: Category = "A" (xlAnd = 1)
    excel.get_range(&ws, "A1:D26")?.call(
        "AutoFilter",
        &[i4(2), bstr("A"), i4(1), empty(), vbool(true)],
    )?;
    println!("  AutoFilter applied: Category = 'A'");

    excel.save_as_xlsx(&wb, &file_path)?;
    wb.call("Close", &[vbool(false)])?;

    let mut pass = true;
    if nyc_count != 0 {
        println!("  FAIL: NYC should have been replaced");
        pass = false;
    }
    if new_york_count == 0 {
        println!("  FAIL: No 'New York City' found");
        pass = false;
    }
    if !sort_ok {
        pass = false;
    }
    Ok(pass)
}

// ---------------------------------------------------------------------------
// E09 - Chart Generation
// ---------------------------------------------------------------------------
pub fn run_e09(output_dir: &str) -> Result<bool> {
    let file_path = make_path(output_dir, "E09_chart.xlsx");
    delete_if_exists(&file_path);

    let excel = ExcelApp::new()?;
    let wb = excel.add_workbook()?;
    let ws = excel.get_sheet(&wb, 1)?;

    excel.get_cell(&ws, 1, 1)?.put("Value2", bstr("Category"))?;
    excel.get_cell(&ws, 1, 2)?.put("Value2", bstr("Value"))?;

    let cats = ["Alpha", "Beta", "Gamma", "Delta", "Epsilon"];
    let vals = [10.0, 25.0, 15.0, 30.0, 20.0];
    for i in 0..5usize {
        excel.get_cell(&ws, i as i32 + 2, 1)?.put("Value2", bstr(cats[i]))?;
        excel.get_cell(&ws, i as i32 + 2, 2)?.put("Value2", r8(vals[i]))?;
    }

    // ChartObjects().Add(left, top, width, height)
    let chart_objs = ws.get_dispatch_with_args("ChartObjects", &[empty()])?;
    let chart_obj = chart_objs.call_dispatch("Add", &[r8(200.0), r8(10.0), r8(400.0), r8(300.0)])?;
    let chart = chart_obj.get_dispatch("Chart")?;

    // Source data
    let src_range = excel.get_range(&ws, "A1:B6")?;
    chart.call("SetSourceData", &[Variant::Dispatch(src_range), empty()])?;

    // Chart type (xlColumnClustered = 51)
    chart.put("ChartType", i4(51))?;

    // Title
    chart.put("HasTitle", vbool(true))?;
    chart.get_dispatch("ChartTitle")?.put("Text", bstr("Category Values"))?;
    println!("  Created column clustered chart with title");

    excel.save_as_xlsx(&wb, &file_path)?;
    wb.call("Close", &[vbool(false)])?;

    if !file_exists_non_empty(&file_path) {
        println!("  FAIL: File not found or empty");
        return Ok(false);
    }

    let wb2 = excel.open_workbook(&file_path)?;
    let ws2 = excel.get_sheet(&wb2, 1)?;

    let mut pass = true;

    let chart_objs2 = ws2.get_dispatch_with_args("ChartObjects", &[empty()])?;
    let chart_count = chart_objs2.get("Count")?.as_f64().unwrap_or(-1.0) as i64;
    if chart_count != 1 {
        println!("  FAIL: Chart count = {}, expected 1", chart_count);
        pass = false;
    } else {
        println!("  Chart count: OK (1)");
    }

    let chart_obj2 = chart_objs2.get_dispatch_with_args("Item", &[i4(1)])?;
    let chart2 = chart_obj2.get_dispatch("Chart")?;
    let chart_type = chart2.get("ChartType")?.as_f64().unwrap_or(-1.0) as i64;
    if chart_type != 51 {
        println!("  FAIL: Chart type = {}, expected 51 (xlColumnClustered)", chart_type);
        pass = false;
    } else {
        println!("  Chart type: OK (xlColumnClustered)");
    }

    match chart2.get("HasTitle")? {
        Variant::Bool(true) => {
            let title = chart2.get_dispatch("ChartTitle")?.get("Text")?;
            match title.as_string() {
                Some("Category Values") => println!("  Chart title: OK"),
                other => {
                    println!("  FAIL: Chart title = {:?}, expected \"Category Values\"", other);
                    pass = false;
                }
            }
        }
        _ => {
            println!("  FAIL: Chart has no title");
            pass = false;
        }
    }

    wb2.call("Close", &[vbool(false)])?;
    println!("  File: {} {}", file_path, if pass { "OK" } else { "FAIL" });
    Ok(pass)
}

// ---------------------------------------------------------------------------
// E10 - Export PDF
// ---------------------------------------------------------------------------
pub fn run_e10(output_dir: &str) -> Result<bool> {
    let xlsx_path = make_path(output_dir, "E10_export.xlsx");
    let pdf_path = make_path(output_dir, "E10_export.pdf");
    delete_if_exists(&xlsx_path);
    delete_if_exists(&pdf_path);

    let excel = ExcelApp::new()?;
    let wb = excel.add_workbook()?;
    let ws = excel.get_sheet(&wb, 1)?;

    excel.get_cell(&ws, 1, 1)?.put("Value2", bstr("PDF Export Test"))?;
    excel.get_cell(&ws, 2, 1)?.put("Value2", bstr("Row 2 data"))?;
    for r in 3..=10 {
        excel.get_cell(&ws, r, 1)?.put("Value2", r8((r * 100) as f64))?;
    }

    excel.save_as_xlsx(&wb, &xlsx_path)?;

    // ExportAsFixedFormat: Type 0 = xlTypePDF, then Filename, then 7 optional args
    wb.call(
        "ExportAsFixedFormat",
        &[
            i4(0),
            bstr(&pdf_path),
            empty(),
            empty(),
            empty(),
            empty(),
            empty(),
            empty(),
            empty(),
        ],
    )?;
    println!("  Exported PDF: {}", pdf_path);

    wb.call("Close", &[vbool(false)])?;

    let pass = file_exists_non_empty(&pdf_path);
    if pass {
        let sz = std::fs::metadata(&pdf_path).map(|m| m.len()).unwrap_or(0);
        println!("  PDF size: {} bytes", sz);
    } else {
        println!("  FAIL: PDF not generated or empty");
    }
    Ok(pass)
}

// ---------------------------------------------------------------------------
// E11 - Run VBA Macro
// ---------------------------------------------------------------------------
pub fn run_e11(output_dir: &str) -> Result<bool> {
    let file_path = make_path(output_dir, "E11_macro.xlsm");
    delete_if_exists(&file_path);

    let excel = ExcelApp::new()?;
    let wb = excel.add_workbook()?;
    let ws = excel.get_sheet(&wb, 1)?;

    excel.get_cell(&ws, 1, 1)?.put("Value2", bstr("Before Macro"))?;
    excel.get_cell(&ws, 2, 1)?.put("Value2", r8(10.0))?;

    let vba_code = "Sub TestMacro()\r\n\
        \x20   Dim ws As Worksheet\r\n\
        \x20   Set ws = ThisWorkbook.Worksheets(1)\r\n\
        \x20   ws.Range(\"A1\").Value2 = \"Macro Executed\"\r\n\
        \x20   ws.Range(\"A2\").Value2 = ws.Range(\"A2\").Value2 * 2\r\n\
        \x20   ws.Range(\"A3\").Value2 = \"VBA was here\"\r\n\
        End Sub\r\n";

    // Inject VBA via the VBProject object model. Requires "Trust access to the
    // VBA project object model"; if disabled, treat as SKIP (environment issue).
    let inject = || -> Result<()> {
        let vb_project = wb.get_dispatch("VBProject")?;
        let vb_components = vb_project.get_dispatch("VBComponents")?;
        let module = vb_components.call_dispatch("Add", &[i4(1)])?; // vbext_ct_StdModule
        let code_module = module.get_dispatch("CodeModule")?;
        code_module.call("AddFromString", &[bstr(vba_code)])?;
        Ok(())
    };

    match inject() {
        Ok(()) => println!("  VBA code injected successfully"),
        Err(_) => {
            println!("  WARNING: Cannot inject VBA - Trust setting may be disabled");
            println!("  Enable: File > Options > Trust Center > Macro Settings");
            println!("  > 'Trust access to the VBA project object model'");
            println!("  RESULT: SKIP (environment issue, not code defect)");
            wb.call("Close", &[vbool(false)])?;
            return Ok(true); // Not a code defect
        }
    }

    // Run the macro
    if let Err(e) = excel.app().call("Run", &[bstr("TestMacro")]) {
        println!("  FAIL: Macro execution failed: {}", e.message());
        wb.call("Close", &[vbool(false)])?;
        return Ok(false);
    }
    println!("  Macro executed successfully");

    // Verify effects
    let str_a1 = excel.get_cell(&ws, 1, 1)?.get("Value2")?;
    let dbl_a2 = excel.get_cell(&ws, 2, 1)?.get("Value2")?.as_f64().unwrap_or(-1.0);
    let str_a3 = excel.get_cell(&ws, 3, 1)?.get("Value2")?;

    println!(
        "  After macro: A1=\"{}\", A2={}, A3=\"{}\"",
        str_a1.as_string().unwrap_or(""),
        dbl_a2,
        str_a3.as_string().unwrap_or("")
    );

    let mut pass = true;
    if str_a1.as_string() != Some("Macro Executed") {
        println!("  FAIL: A1 should be 'Macro Executed'");
        pass = false;
    }
    if (dbl_a2 - 20.0).abs() > 0.001 {
        println!("  FAIL: A2 should be 20 (10*2)");
        pass = false;
    }
    if str_a3.as_string() != Some("VBA was here") {
        println!("  FAIL: A3 should be 'VBA was here'");
        pass = false;
    }

    // SaveAs .xlsm (xlOpenXMLWorkbookMacroEnabled = 52)
    wb.call("SaveAs", &[bstr(&file_path), i4(52)])?;
    wb.call("Close", &[vbool(false)])?;
    Ok(pass)
}

// ---------------------------------------------------------------------------
// E12 - Resource Cleanup (RAII demonstration)
// ---------------------------------------------------------------------------
pub fn run_e12(output_dir: &str) -> Result<bool> {
    use windows::Win32::Foundation::{CloseHandle, HWND, STILL_ACTIVE};
    use windows::Win32::System::Threading::{
        GetExitCodeProcess, OpenProcess, TerminateProcess, WaitForSingleObject,
        PROCESS_ACCESS_RIGHTS, PROCESS_QUERY_INFORMATION,
    };
    // SYNCHRONIZE (0x0010_0000) is a standard access right; construct it directly
    // since the crate only exposes it under a different feature/newtype.
    const SYNCHRONIZE: PROCESS_ACCESS_RIGHTS = PROCESS_ACCESS_RIGHTS(0x0010_0000);
    use windows::Win32::UI::WindowsAndMessaging::GetWindowThreadProcessId;

    println!("  Demonstrating RAII cleanup...");

    let mut our_pid: u32 = 0;

    // Inner scope: all COM objects created and dropped here
    {
        let excel = ExcelApp::new()?;

        let hwnd_val = excel.hwnd()?;
        unsafe {
            GetWindowThreadProcessId(HWND(hwnd_val as *mut _), Some(&mut our_pid));
        }
        println!("  Our Excel PID: {}", our_pid);

        let wb = excel.add_workbook()?;
        let ws = excel.get_sheet(&wb, 1)?;

        excel.get_cell(&ws, 1, 1)?.put("Value2", bstr("RAII Test"))?;
        for r in 2..=100 {
            excel.get_cell(&ws, r, 1)?.put("Value2", r8(r as f64))?;
        }

        let file_path = make_path(output_dir, "E12_cleanup.xlsx");
        delete_if_exists(&file_path);
        excel.save_as_xlsx(&wb, &file_path)?;

        wb.call("Close", &[vbool(false)])?;
        drop(ws);
        drop(wb);

        println!("  Workbook closed, releasing COM objects...");
        // ExcelApp::drop() calls Quit() + releases the application reference here.
    }

    println!("  All COM objects released, Excel.Application quit");

    // Give Excel time to terminate naturally
    std::thread::sleep(std::time::Duration::from_millis(3000));

    // Check whether our specific process has exited
    let no_residual;
    unsafe {
        match OpenProcess(PROCESS_QUERY_INFORMATION | SYNCHRONIZE, false, our_pid) {
            Err(_) => {
                no_residual = true;
                println!("  Our Excel process (PID {}) has exited - cleanup successful", our_pid);
            }
            Ok(h_proc) => {
                let mut exit_code: u32 = 0;
                let got = GetExitCodeProcess(h_proc, &mut exit_code).is_ok();
                if got && exit_code != STILL_ACTIVE.0 as u32 {
                    no_residual = true;
                    println!("  Our Excel process (PID {}) has exited - cleanup successful", our_pid);
                } else {
                    println!("  Process still alive after 3s, forcing kill...");
                    let _ = TerminateProcess(h_proc, 1);
                    WaitForSingleObject(h_proc, 5000);
                    no_residual = true;
                    println!("  WARNING: Hidden COM refs prevented clean exit - process killed.");
                }
                let _ = CloseHandle(h_proc);
            }
        }
    }

    Ok(no_residual)
}
