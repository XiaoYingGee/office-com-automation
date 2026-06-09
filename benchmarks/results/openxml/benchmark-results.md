# Excel Automation Benchmark Results

> Date: 2026-06-09 13:18
> Backends: openxml

## Latency (ms)

| Benchmark | openxml |
|-----------|------:|
| Cell Write (5 cells) | 1647ms |
| Bulk Write (100x10) | 352ms |
| Read Cell | 317ms |
| Clear Range (1000 cells) | 336ms |
| Inspect Workbook | 318ms |
| Batch 10 Writes | 3181ms |
| Format 5 Cells (bold+size+color) | 1621ms |
| Insert 5 Rows | 311ms |
| Sheet Add+Rename+Delete | 918ms |

## Notes

- **pywin32**: Python → cross-process COM IPC (one call per property)
- **vba**: Python → Application.Run → in-process VBA (zero IPC for object model)
- **rust**: Standalone exe, Excel COM, one process spawn per operation
- **openxml**: Standalone exe, direct .xlsx file manipulation, no Excel process

### Architecture

```
pywin32:  Python ──COM IPC (per property)──→ Excel process
vba:      Python ──App.Run (1x)──→ Excel process内 VBA execution
rust:     [spawn] ──COM──→ Excel process ──close ── (per op)
openxml:  [spawn] ──file I/O──→ .xlsx (no Excel)
```