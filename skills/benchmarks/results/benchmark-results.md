# Excel Automation Benchmark Results

> Date: 2026-06-09 12:59
> Backends: pywin32, vba

## Latency (ms)

| Benchmark | pywin32 | vba |
|-----------|------:|------:|
| Cell Write (5 cells) | 509ms | 408ms |
| Bulk Write (100x10) | 37ms | 28ms |
| Read Cell | 36ms | 5ms |
| Clear Range (1000 cells) | 20ms | 21ms |
| Inspect Workbook | 27ms | 24ms |
| Batch 10 Writes | 135ms | 9ms |
| Format 5 Cells (bold+size+color) | 183ms | 36ms |
| Insert 5 Rows | 38ms | 20ms |
| Sheet Add+Rename+Delete | 51ms | 15ms |

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