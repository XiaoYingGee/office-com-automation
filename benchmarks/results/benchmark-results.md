# Excel Automation Benchmark Results

> Date: 2026-06-09 14:49
> Sizes: empty, medium, large

## Size: empty  (latency, lower = faster)

| Benchmark | pywin32 | vba |
|-----------|------:|------:|
| Open Workbook | N/A (no file) | N/A (no file) |
| Cell Write (5 cells) | 488 ms | 422 ms |
| Bulk Write (100x10) | 33 ms | 31 ms |
| Read Cell | 25 ms | 7 ms |
| Clear Range (1000 cells) | 27 ms | 20 ms |
| Inspect Workbook | 28 ms | 13 ms |
| Batch 10 Writes | 139 ms | 6 ms |
| Format 5 Cells (bold+size+color) | 172 ms | 39 ms |
| Insert 5 Rows | 38 ms | 15 ms |
| Sheet Add+Rename+Delete | 51 ms | 20 ms |
| Merge 5 Ranges | 105 ms | 6 ms |

## Size: medium  (latency, lower = faster)

| Benchmark | pywin32 | vba |
|-----------|------:|------:|
| Open Workbook | 751 ms | 889 ms |
| Cell Write (5 cells) | 424 ms | 551 ms |
| Bulk Write (100x10) | 26 ms | 17 ms |
| Read Cell | 13 ms | 4 ms |
| Clear Range (1000 cells) | 78 ms | 6 ms |
| Inspect Workbook | 84 ms | 4 ms |
| Batch 10 Writes | 214 ms | 6 ms |
| Format 5 Cells (bold+size+color) | 123 ms | 26 ms |
| Insert 5 Rows | 46 ms | 13 ms |
| Sheet Add+Rename+Delete | 52 ms | 15 ms |
| Merge 5 Ranges | 204 ms | 4 ms |

## Size: large  (latency, lower = faster)

| Benchmark | pywin32 | vba |
|-----------|------:|------:|
| Open Workbook | 18532 ms | 18255 ms |
| Cell Write (5 cells) | 452 ms | 527 ms |
| Bulk Write (100x10) | 23 ms | 21 ms |
| Read Cell | 16 ms | 2 ms |
| Clear Range (1000 cells) | 105 ms | 11 ms |
| Inspect Workbook | 504 ms | 121 ms |
| Batch 10 Writes | 147 ms | 28 ms |
| Format 5 Cells (bold+size+color) | 173 ms | 84 ms |
| Insert 5 Rows | 150 ms | 117 ms |
| Sheet Add+Rename+Delete | 65 ms | 21 ms |
| Merge 5 Ranges | 233 ms | 12 ms |

## Notes

- **pywin32**: Python -> cross-process COM IPC (one call per property), persistent session.
- **vba**: Python -> Application.Run -> in-process VBA (zero IPC), persistent session.
- **rust / cpp**: standalone COM exe, one Excel process spawn per op (stateless).
- **openxml**: standalone .NET exe, whole-file read+write per op, no Excel (stateless).
- Stateless backends are excluded at the `large` tier: re-opening a ~145 MB file per op is architecturally infeasible (exceeds timeout / takes hours). Shown as N/A.
- `B0 Open` is N/A for the empty tier (no file to open).