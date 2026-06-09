# Excel Automation Benchmark Results

> Date: 2026-06-09
> Backends measured: **pywin32, vba, openxml** (3/4)
> Rust: not measured — stateless design spawns a fresh Excel process per op, prohibitively slow for this micro-op suite (see Notes).
> Rounds: 1 warmup + 5 measured. pywin32/vba from one run; openxml from a separate run (same machine).

## Latency — total ms per test (lower = faster)

| # | Benchmark | pywin32 | vba | openxml |
|---|-----------|------:|------:|------:|
| B1 | Cell Write (5 cells) | 509 | **408** | 1647 |
| B2 | Bulk Write (100×10) | 37 | **28** | 352 |
| B3 | Read Cell | 36 | **5** | 317 |
| B4 | Clear Range (1000 cells) | **20** | 21 | 336 |
| B5 | Inspect Workbook | 27 | **24** | 318 |
| B6 | Batch 10 Writes | 135 | **9** | 3181 |
| B7 | Format 5 Cells (bold+size+color) | 183 | **36** | 1621 |
| B8 | Insert 5 Rows | 38 | **20** | 311 |
| B9 | Sheet Add+Rename+Delete | 51 | **15** | 918 |

**Measured ranking: 🥇 VBA → 🥈 pywin32 → 🥉 OpenXML** (Rust not run)

## Notes

- **pywin32**: Python → cross-process COM IPC (one call per property)
- **vba**: Python → `Application.Run` → in-process VBA (zero IPC for object model)
- **openxml**: standalone .NET exe, direct .xlsx file manipulation, no Excel process
- **rust**: standalone exe, Excel COM, one process spawn per op — *not measured*

### Finding 1 — VBA dominates IPC-heavy ops (matches expectation)

VBA wins 8/9 tests. The gap is widest exactly where cross-process COM IPC dominates:
B6 Batch 10 Writes (135 → 9, **15×**), B7 Format (183 → 36, **5×**), B3 Read (36 → 5, **7×**).
This confirms the architectural prediction: VBA executes in-process with zero per-property
IPC, while pywin32 pays one COM round-trip per property access.

### Finding 2 — OpenXML is the SLOWEST here, contradicting the original prediction

The handoff originally predicted OpenXML as 🥉 / "fast file I/O". Measured reality: OpenXML is
the slowest backend on **every** test, by a large margin (B6 = 3181ms, ~350× VBA).

Root cause: the OpenXML backend is a **standalone exe invoked per op**, and each invocation
pays *process startup + read the entire .xlsx + mutate + write the whole file back*. For this
micro-op suite, that fixed overhead (~300ms floor, visible across B3/B4/B5/B8 all ≈310–340ms)
dwarfs the savings of not launching Excel. B6 (10 writes) is worst because it repeats the
full-file rewrite 10×.

OpenXML's advantage only materializes for **large, batched, one-shot** transformations where
Excel never needs to be live — not for small high-frequency ops like this benchmark.

### Architecture

```
pywin32:  Python ──COM IPC (per property)──→ Excel process
vba:      Python ──App.Run (1x)──→ in-process VBA execution
openxml:  [spawn exe] ──read+write whole .xlsx──→ file (no Excel)   ← per-op overhead
rust:     [spawn exe] ──COM──→ Excel process ──close── (per op)     ← not measured
```
