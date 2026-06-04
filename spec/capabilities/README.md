# Capability System — Orientation

This directory contains the fine-grained **capability catalog** for the Excel automation harness.
A capability is a single testable property of one Excel operation (e.g. "write a boolean to a cell
and read it back"). The catalog drives automated cross-backend verification via `capctl verify`.

---

## What the Catalog Is

Rather than coarse task-level pass/fail, the catalog tracks individual, single-property
capabilities. Each entry records:

- the op name and the specific parameter variant being tested,
- a complete verification recipe (`verify.setup[]` → `verify.action` → `verify.assert.read`),
- per-backend support status (✅ / ⚠️ / ❌ / ⬜).

This makes it possible to say, for example, "OpenXML passes `CELL-WRITE-STRING` but ❌ on
`CELL-WRITE-FORMULA` (no formula recalculation at file level)" rather than just "partially
supports cell I/O".

---

## File Layout

```
spec/capabilities/
├── README.md                  ← you are here
├── contract.md                ← semantics and JSON shapes (OpRequest / OpResponse / error model)
├── schema/
│   ├── op-request.schema.json ← JSON Schema for OpRequest
│   ├── error.schema.json      ← JSON Schema for ExcelError
│   └── capability.schema.json ← JSON Schema for capability catalog entries
├── catalog/
│   └── 01-cell-range-io.md   ← Cell & Range I/O domain (9 capabilities; current first slice)
└── support-matrix.md          ← generated cross-backend support matrix (do not edit by hand)
```

More domain files will be added under `catalog/` as new capability domains are specified
(formatting, row/column structure, multi-sheet, data operations, charts, PDF export, VBA macros).

---

## Support Legend

| Symbol | Meaning |
|--------|---------|
| ✅     | Full — action succeeded and read-back matched expected value |
| ⚠️     | Partial / lossy — action ran but read-back differed from expected |
| ❌     | Unsupported / errored — backend returned an error, or reference read failed |
| ⬜     | Untested — `capctl verify` has not been run for this backend/capability |

---

## Running `capctl verify`

To regenerate `support-matrix.md` for a backend:

```
capctl verify \
  --backend     <name>              \   # vba | cpp | rust | openxml
  --backend-cmd  <path/to/backend>  \
  --reference-cmd <path/to/rust-excel-ops> \
  --catalog     spec/capabilities/catalog \
  --out         spec/capabilities/support-matrix.md
```

The reference reader (`--reference-cmd`) must always be the Rust `excel-ops` binary so that all
backends are compared against the same consistent reader.

See [`docs/backend-protocol.md`](../../docs/backend-protocol.md) for the full protocol a backend
executable must implement, including op shapes, error categories, and stateless semantics.

---

## Current Status

**Rust reference backend only. Cell I/O domain only.**

- 1 backend implemented: Rust COM (`languages/rust/`)
- 1 domain implemented: Cell & Range I/O (`catalog/01-cell-range-io.md`) — 9 capabilities
- 3 other backends planned: VBA, C++, OpenXML (.NET, file-level) — not yet implemented
- 7 other domains planned: formatting, row/column structure, multi-sheet, data operations,
  charts, PDF export, VBA macros — not yet specified/implemented

---

## Further Reading

- [`contract.md`](contract.md) — full op contract, error model, and `batch.apply` semantics
- [`docs/backend-protocol.md`](../../docs/backend-protocol.md) — implementer guide for new backends
- [`docs/superpowers/specs/2026-06-04-excel-capability-catalog-design.md`](../../docs/superpowers/specs/2026-06-04-excel-capability-catalog-design.md) — design rationale and domain list
