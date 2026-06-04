# Rust Implementation Handoff Document

> **Date:** 2026-06-04
> **Status:** Phase 2 complete (com.rs compiles), Phase 3-5 pending
> **Goal:** Implement Rust E01-E12 Excel COM automation tasks, achieving 12/12 PASS

---

## Project Overview

This repo (`office-com-automation`) implements 12 Excel COM automation tasks (E01-E12) in 3 languages for comparison: **VBA, C++, Rust**. VBA and C++ are both 12/12 PASS. Rust is the new addition replacing C#.

- **Repo:** `D:\Workspace\AI\office-com-automation`
- **Rust project:** `languages/rust/`
- **C++ reference:** `languages/cpp/src/` (primary porting source)
- **VBA reference:** `languages/vba/src/`
- **Task spec:** `spec/excel-tasks.md`

## Current Git State

All changes are **uncommitted** on the initial branch. Key changes already made:

- Deleted `languages/{csharp,python,powershell,jscript}/` (narrowed to 3 languages)
- Updated `README.md`, `spec/capability-matrix.md`, `spec/excel-tasks.md`, `docs/pk-framework.md`, `docs/roadmap.md` (6→3 languages)
- Added `target/` to `.gitignore` for Rust
- Created full `languages/rust/` project structure

## Rust Project Structure

```
languages/rust/
├── Cargo.toml          # windows crate v0.61, edition 2021
└── src/
    ├── main.rs         # STUB - needs CLI runner implementation
    ├── com.rs          # DONE - compiles, IDispatch/Variant/SafeArray wrappers
    ├── excel.rs        # STUB - needs ExcelApp RAII wrapper
    └── tasks.rs        # STUB - needs E01-E12 implementations
```

### Cargo.toml Dependencies

```toml
[dependencies.windows]
version = "0.61"
features = [
    "Win32_System_Com",
    "Win32_System_Ole",
    "Win32_System_Variant",
    "Win32_Foundation",
    "Win32_UI_WindowsAndMessaging",
    "Win32_System_Threading",
]
```

## What's Done: `com.rs` (COMPLETE)

Full IDispatch late-binding wrapper. Key types:

| Type | Purpose |
|------|---------|
| `ComGuard` | RAII for `CoInitializeEx`/`CoUninitialize` |
| `Dispatch` | Wraps `IDispatch` — `get()`, `put()`, `call()`, `get_dispatch()`, `get_with_args()`, `call_dispatch()`, `get_dispatch_with_args()` |
| `Variant` | Safe enum: `Empty`, `I4(i32)`, `R8(f64)`, `Bool(bool)`, `Bstr(String)`, `Dispatch(Dispatch)`, `Unknown` |
| `SafeArray2D` | Build 2D `SAFEARRAY(VT_VARIANT)` for bulk Range.Value2 writes |
| Helper fns | `empty()`, `i4()`, `r8()`, `bstr()`, `vbool()`, `cell_address()`, `range_address()` |

### Critical Implementation Notes

1. **ManuallyDrop in Rust 1.94:** VARIANT union field access requires `let v00 = &mut *v.Anonymous.Anonymous;` — the first level (`v.Anonymous`) is a plain union, the second level (`.Anonymous`) is `ManuallyDrop<VARIANT_0_0>` requiring explicit deref.

2. **SafeArray functions** are in `Win32::System::Ole`, NOT `Win32::System::Com`.

3. **VARIANT_BOOL** is in `Win32::Foundation`.

4. **CoInitializeEx** returns HRESULT — use `.ok()?` to convert to Result.

5. **put()** must set `DISPID_PROPERTYPUT` in `rgdispidNamedArgs`.

6. **call()** must reverse argument order in `rgvarg` (COM convention).

7. **Parameterized properties** (e.g., `Range("A1")`) use `get_with_args()` with `DISPATCH_PROPERTYGET | DISPATCH_METHOD`.

## What's Pending

### Phase 3: `excel.rs` — ExcelApp RAII Wrapper

Port from C++ `excel_com.h` / `excel_com.cpp`. Structure:

```rust
pub struct ExcelApp {
    app: Dispatch,  // Excel.Application
}
```

**Constructor:**
- `Dispatch::create("Excel.Application")`
- Set `Visible` = false, `DisplayAlerts` = false, `ScreenUpdating` = false

**Drop:**
- Restore `DisplayAlerts` = true, `ScreenUpdating` = true
- Call `Quit()`

**Methods to implement:**
- `add_workbook() -> Result<Dispatch>` — `app.get_dispatch("Workbooks")?.call_dispatch("Add", &[empty()])`
- `open_workbook(path) -> Result<Dispatch>` — `workbooks.call_dispatch("Open", &[bstr(path)])`
- `get_sheet(wb, index) -> Result<Dispatch>` — `wb.get_dispatch("Worksheets")?.get_dispatch_with_args("Item", &[i4(index)])`
- `get_sheet_by_name(wb, name) -> Result<Dispatch>` — same with `bstr(name)`
- `add_sheet(wb) -> Result<Dispatch>` — `worksheets.call_dispatch("Add", &[empty(), empty(), empty(), empty()])`
- `get_range(ws, address) -> Result<Dispatch>` — `ws.get_dispatch_with_args("Range", &[bstr(address)])`
- `get_cell(ws, row, col) -> Result<Dispatch>` — builds "A1" address, calls `get_range`
- `bulk_write(ws, start_row, start_col, rows, cols, sa_variant)` — set Range.Value2 to SAFEARRAY
- `hwnd() -> Result<i64>` — `app.get("Hwnd")?.as_f64()` cast to i64
- `save_as_xlsx(wb, path)` — `wb.call("SaveAs", &[bstr(path), i4(51)])` (51 = xlOpenXMLWorkbook)

### Phase 4: `tasks.rs` — E01-E12

Direct port of C++ `tasks.cpp` (~1093 lines). Each task: `pub fn run_eXX(output_dir: &str) -> Result<bool>`.

| Task | Description | Key Notes |
|------|-------------|-----------|
| E01 | Workbook Lifecycle | Create, save, close, reopen, verify cell value |
| E02 | Cell Read/Write | String, number, boolean, date serial. Read back and verify types |
| E03 | Bulk Range Write | `SafeArray2D::fill_sequential()` vs cell-by-cell, timing comparison |
| E04 | Formula & Recalc | Manual calc mode (`put("Calculation", i4(-4135))`), write formulas, Calculate, verify |
| E05 | Cell Formatting | Font bold/size/color, interior color, number format, alignment, merge |
| E06 | Row/Column Structure | Insert/delete rows, RowHeight, ColumnWidth, AutoFit |
| E07 | Multi-Worksheet | Rename sheet, cross-sheet formula, add/delete temp sheet, verify count |
| E08 | Data Operations | Replace (8 args), Sort (15 args — pad with `empty()`), AutoFilter |
| E09 | Chart Generation | `ChartObjects` → `Add` → `Chart` → SetSourceData/ChartType/Title |
| E10 | Export PDF | `ExportAsFixedFormat` with type=0 (xlTypePDF) |
| E11 | Run VBA Macro | VBProject → VBComponents → Add(1) → CodeModule → AddFromString. Then `app.call("Run", ...)` |
| E12 | Resource Cleanup | Inner scope for RAII Drop. `GetWindowThreadProcessId` for PID, 3s wait, force kill if needed |

**E08 Sort COM constants:**
- `xlDescending` = 2, `xlAscending` = 1
- `xlYes` = 1 (header)
- `xlSortColumns` = 1, `xlPinYin` = 1, `xlSortNormal` = 0

**E09 Chart constants:**
- `xlColumnClustered` = 51

**E10 PDF export:**
- `xlTypePDF` = 0

**E11 VBA module type:**
- `vbext_ct_StdModule` = 1
- Save as xlsm: format code = 52 (`xlOpenXMLWorkbookMacroEnabled`)

**E12 Win32 APIs needed:**
- `GetWindowThreadProcessId` — from `Win32_UI_WindowsAndMessaging` feature
- `OpenProcess`, `GetExitCodeProcess`, `TerminateProcess`, `WaitForSingleObject` — from `Win32_System_Threading`
- `Sleep` — from `Win32_System_Threading`
- `CloseHandle` — from `Win32_Foundation`

### Phase 5: `main.rs` — CLI Runner

Port from C++ `main.cpp`. Structure:

```rust
use std::time::Instant;
use std::path::PathBuf;
use std::fs;
use std::env;

mod com;
mod excel;
mod tasks;

fn main() {
    // 1. Determine output dir: <exe_dir>/out or <cwd>/out
    // 2. Parse args: no args = all tasks, else specific E01-E12
    // 3. Init COM once: ComGuard::new()
    // 4. Run tasks with timing (Instant::now / elapsed)
    // 5. Print summary table: Task / Result / Time
    // 6. Exit code: 0 if all pass, 1 otherwise
}
```

Task registry pattern:
```rust
type TaskFn = fn(&str) -> Result<bool>;
struct TaskEntry { name: &'static str, func: TaskFn }
let tasks = vec![
    TaskEntry { name: "E01", func: tasks::run_e01 },
    // ... E02-E12
];
```

## Build & Run Commands

```bash
cd D:/Workspace/AI/office-com-automation/languages/rust
cargo build --release
cargo run --release           # run all tasks
cargo run --release -- E01    # run specific task
```

## Verification Criteria

1. `cargo build --release` — no errors
2. `cargo run --release` — 12/12 PASS
3. Output files in `languages/rust/src/out/` (or `target/release/out/`) are valid xlsx/pdf
4. Summary table matches C++ format

## Expanded Scope (After Rust 12/12 PASS)

User has requested a much larger follow-up phase:

1. **Comprehensive Excel API wrapping** — merge cells, fonts, colors, formats, borders, backgrounds, formulas, charts, and ALL COM-supported features (popular + niche)
2. **Boundary condition tests** — success + failure cases with standardized error returns
3. **MCP/skill integration** — each language produces AI-callable tools
4. **Documentation** — summarize capabilities, pitfalls, push to GitHub

This is NOT in scope for the current handoff. Complete Rust 12/12 PASS first.

## File References

| File | Purpose | Status |
|------|---------|--------|
| `languages/rust/src/com.rs` | COM wrappers | DONE, compiles |
| `languages/rust/src/excel.rs` | ExcelApp RAII | STUB (`pub struct ExcelApp;`) |
| `languages/rust/src/tasks.rs` | E01-E12 | STUB (`pub fn placeholder() {}`) |
| `languages/rust/src/main.rs` | CLI runner | STUB |
| `languages/rust/Cargo.toml` | Dependencies | DONE |
| `languages/cpp/src/tasks.cpp` | C++ reference (1093 lines) | Reference for porting |
| `languages/cpp/src/excel_com.cpp` | C++ ExcelApp reference | Reference for porting |
| `languages/cpp/src/excel_com.h` | C++ headers + HiResTimer | Reference for porting |
| `languages/cpp/src/main.cpp` | C++ CLI runner reference | Reference for porting |
| `spec/excel-tasks.md` | Task specifications | Reference |
| `.claude/plans/memoized-baking-wirth.md` | Implementation plan | Reference |
