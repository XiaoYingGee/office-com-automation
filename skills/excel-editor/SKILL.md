---
name: excel-editor
description: "Edit Excel files via COM automation on a local Windows machine. Two backends: pywin32 (baseline) and VBA (in-process, 10-20x faster for reads). Triggers on: excel editing, xlsx automation, local excel, spreadsheet editor, excel com."
---

# Excel COM Editor Skill

## Overview

Local Excel automation via COM. Two execution backends:

- **pywin32** (default): Direct Python COM calls. Zero setup, works immediately.
- **vba**: In-process VBA macros via `Application.Run`. 10-20x faster for inspect/reads (zero IPC). Requires "Trust access to VBA project object model" enabled in Excel Trust Center.

## Usage

### Inspect workbook structure

```bash
python scripts/excel_editor.py workbook.xlsx --inspect
```

### Inspect sheet data

```bash
python scripts/excel_editor.py workbook.xlsx --inspect-sheet Sheet1
python scripts/excel_editor.py workbook.xlsx --inspect-sheet Sheet1 --max-rows 100
```

### Execute JSON actions

```bash
python scripts/excel_editor.py workbook.xlsx --exec-actions '[{"action":"write_cell","sheet":"Sheet1","cell":"A1","value":42}]'
python scripts/excel_editor.py workbook.xlsx --exec-actions actions.json
python scripts/excel_editor.py workbook.xlsx --exec-actions actions.json --output result.xlsx
```

### Execute Python script

```bash
python scripts/excel_editor.py workbook.xlsx --exec-script edit.py
```

Script has `excel` (engine instance) and `filepath` injected as globals:

```python
# edit.py
data = excel.inspect()
excel.write_cell("Sheet1", "A1", "Hello")
excel.write_range("Sheet1", "B1:C3", [[1,2],[3,4],[5,6]])
excel.set_format("Sheet1", "A1:C1", bold=True, bg_color=0x00FFFF)
excel.save()
```

### Interactive JSON session

```bash
python scripts/excel_editor.py workbook.xlsx --interactive-actions
```

Commands: `inspect`, `save`, `quit`, or JSON action(s).

### Create new workbook

```bash
python scripts/excel_editor.py new_file.xlsx --create
```

### Use VBA backend

```bash
python scripts/excel_editor.py workbook.xlsx --inspect --backend vba
python scripts/excel_editor.py workbook.xlsx --exec-actions actions.json --backend vba
```

### Options

| Option | Description |
|--------|-------------|
| `--backend pywin32\|vba` | Execution backend |
| `--headed` | Show Excel window |
| `--output / -o` | Save to this path |
| `--max-rows` | Max rows for inspect-sheet (default 50) |
| `--max-cols` | Max cols for inspect-sheet (default 26) |

## Supported Actions

### Cell I/O

| Action | Fields | Description |
|--------|--------|-------------|
| `write_cell` | `sheet, cell, value, kind?` | Write value. kind: auto/formula |
| `read_cell` | `sheet, cell, property?` | Read value/formula/text |
| `write_range` | `sheet, range, values` | Bulk write 2D array |
| `read_range` | `sheet, range` | Bulk read 2D array |
| `clear_range` | `sheet, range, mode?` | Clear contents or all |

### Formatting

| Action | Fields | Description |
|--------|--------|-------------|
| `set_format` | `sheet, range, bold?, italic?, font_size?, font_name?, font_color?, bg_color?, number_format?, h_align?, v_align?, merge?, wrap_text?` | Apply formatting |

### Row/Column Structure

| Action | Fields | Description |
|--------|--------|-------------|
| `insert_rows` | `sheet, row, count?` | Insert rows |
| `delete_rows` | `sheet, row, count?` | Delete rows |
| `insert_cols` | `sheet, col, count?` | Insert columns |
| `delete_cols` | `sheet, col, count?` | Delete columns |
| `autofit_columns` | `sheet, range?` | Auto-fit column width |

### Worksheet Management

| Action | Fields | Description |
|--------|--------|-------------|
| `add_sheet` | `name?, after?` | Add new sheet |
| `rename_sheet` | `sheet, new_name` | Rename sheet |
| `delete_sheet` | `sheet` | Delete sheet |

### Data Operations

| Action | Fields | Description |
|--------|--------|-------------|
| `sort_range` | `sheet, range, key_col, order?` | Sort range (asc/desc) |
| `auto_filter` | `sheet, range, field?, criteria?` | Toggle/apply filter |
| `find_replace` | `sheet, range, find, replace` | Find and replace |
| `calculate` | — | Force recalculate |

### Export

| Action | Fields | Description |
|--------|--------|-------------|
| `export_pdf` | `path` | Export workbook to PDF |

## Notes

- **Color format**: COM uses BGR (red = 0x0000FF, blue = 0xFF0000). Python `RGB(r,g,b)` = `r + g*256 + b*65536`.
- **1-based indexing**: All sheet/row/col indices start from 1.
- **VBA Trust**: For `--backend vba`, enable "Trust access to the VBA project object model" in File → Options → Trust Center → Trust Center Settings → Macro Settings.
- **File paths**: Use absolute paths or paths relative to CWD. Avoid OneDrive sync folders.

## Performance Comparison (Expected)

| Operation | pywin32 | VBA | Speedup |
|-----------|---------|-----|---------|
| inspect (large sheet) | ~2-5s | ~100-200ms | 10-20x |
| write_cell (single) | ~50ms | ~15ms | 3x |
| write_range (1000 cells) | ~500ms | ~50ms | 10x |
| read_range (1000 cells) | ~500ms | ~30ms | 15x |
| batch 10 actions | ~500ms | ~100ms | 5x |

Key insight: pywin32 makes one cross-process COM IPC call per property access (~20ms each). VBA runs inside Excel's process with zero IPC overhead.

## Files

| File | Purpose |
|------|---------|
| `scripts/excel_editor.py` | Main entry point + pywin32 engine + VBA backend class |
| `references/ExcelEditorBridge.bas` | VBA bridge module (import into Excel) |
