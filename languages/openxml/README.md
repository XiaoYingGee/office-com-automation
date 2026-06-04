# OpenXML Backend (approach #4 — file-level, no Excel)

A capability backend that manipulates `.xlsx` files **directly via the Open XML SDK**, with **no running Excel process**. It implements the language-neutral op contract in [`docs/backend-protocol.md`](../../docs/backend-protocol.md): read one `OpRequest` JSON from stdin, write one `OpResponse` JSON to stdout.

This is the contrast case in the PK arena: the three COM approaches (VBA/C++/Rust) drive Excel's real engine; OpenXML edits the file format directly — zero Office dependency, but no calculation engine, no chart rendering, no PDF/macro.

## What it implements

Write/setup ops only (the assert-read is always done by the Rust `excel-ops` reference reader, so all approaches are judged by one consistent reader):

| op | notes |
|---|---|
| `cell.write` | kinds: string (inline string), number, bool, **formula** |
| `range.write_bulk` | rectangular 2D values, anchored at the range's top-left |
| `range.clear` | `contents` / `all` (removes the cell) |
| `range.copy_values` | single-cell, same-sheet, type-preserving |

### Honest measurement of formula support
OpenXML has **no calculation engine**. For `cell.write` kind=formula it writes only the formula text (`<f>2+3</f>`) and **does not fabricate a cached result**. Whether `CELL-WRITE-FORMULA` (which asserts the computed value `5`) lands ✅ is then decided by the reference reader: Excel recalculates the cache-less formula when it opens the file, so it currently reads back `5` → ✅. We record reality; we do not fake a cached value.

## Build & run

```bash
dotnet build languages/openxml -c Debug
# emits languages/openxml/bin/Debug/net10.0/ExcelOps.exe
```

Drive it through the orchestrator (from `languages/rust/`):

```bash
cargo run --bin capctl -- verify \
  --backend openxml \
  --backend-cmd ../../languages/openxml/bin/Debug/net10.0/ExcelOps.exe \
  --reference-cmd target/debug/excel-ops.exe \
  --catalog ../../spec/capabilities/catalog \
  --out ../../spec/capabilities/support-matrix.md
```

The matrix merge preserves other backends' columns — running this fills only the **OpenXML** column.

## Tests

```bash
dotnet test languages/openxml
```

24 xUnit tests verify each op by re-reading the produced `.xlsx` **with OpenXML** (no Excel) — proving the files are valid and the values/formulas are written correctly.

## Layout

```
languages/openxml/
├── ExcelOps.csproj      # net10.0 console, DocumentFormat.OpenXml
├── src/
│   ├── Program.cs       # stdin → OpRequest → dispatch → OpResponse → stdout
│   ├── Contract.cs      # OpRequest/OpResponse/error DTOs + JSON options
│   ├── ExcelError.cs    # ErrorCategory + OpException
│   ├── A1.cs            # A1 ref parsing + in-order cell insertion
│   └── Ops.cs           # op implementations over SpreadsheetDocument
└── tests/               # xUnit + OpenXML readback oracle
```

## Status

Verified for the **Cell I/O** domain (9/9 capabilities ✅, via Excel reference reader). Other domains (formatting, structure, charts, …) and the C++/VBA backends are in progress — see [`spec/capabilities/support-matrix.md`](../../spec/capabilities/support-matrix.md).
