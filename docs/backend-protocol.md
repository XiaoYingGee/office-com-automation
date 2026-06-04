# Backend Protocol

This document is the implementation contract every Excel automation backend must satisfy to be
verified against the capability catalog. A backend is any executable (VBA host script, C++ binary,
.NET OpenXML program, Rust binary, …) that speaks the JSON protocol described here.

See also: [`spec/capabilities/contract.md`](../spec/capabilities/contract.md) for the
semantics narrative and [`spec/capabilities/schema/`](../spec/capabilities/schema/) for the
machine-readable JSON schemas.

---

## 1. Process Model

A backend is a **standalone executable** invoked once per operation:

1. capctl spawns the executable.
2. capctl writes **one** `OpRequest` JSON object to the process's **stdin** and closes the pipe.
3. The backend writes **one** `OpResponse` JSON object to **stdout** and exits.
4. Exit code **0** is expected whenever a response was produced — even an error response
   (`ok: false`). Exit code non-zero only on catastrophic failure (e.g. unrecoverable crash before
   any output can be written). capctl treats non-zero exit without valid JSON stdout as a
   spawn-error failure and records ❌ for that capability.
5. **stderr** is inherited (passed through to the terminal); use it for diagnostics.

No persistent state is maintained between calls. Each invocation is independent.

---

## 2. OpRequest Shape

Full schema: [`spec/capabilities/schema/op-request.schema.json`](../spec/capabilities/schema/op-request.schema.json)

```json
{
  "op":     "cell.write",
  "path":   "C:/data/book.xlsx",
  "target": { "sheet": "Sheet1", "range": "B2" },
  "params": { "value": 42, "kind": "number" },
  "save_as": { "path": "C:/data/book-out.xlsx", "format": "xlsx" }
}
```

| Field      | Required | Type   | Description |
|------------|----------|--------|-------------|
| `op`       | yes      | string | Dotted operation name (`cell.write`, `range.read`, …) |
| `path`     | yes      | string | Absolute (or CWD-relative) path to the source workbook |
| `target`   | no       | object | `{ "sheet"?: string, "range"?: string }` — worksheet name and/or A1 range address |
| `params`   | no       | object | Operation-specific parameters; defaults to `{}` |
| `save_as`  | no       | object | `{ "path": string, "format": "xlsx"\|"xlsm"\|"xls"\|"csv" }` — if present, save to this path/format |

---

## 3. OpResponse Shape

Full schema: [`spec/capabilities/schema/error.schema.json`](../spec/capabilities/schema/error.schema.json)

```json
{ "ok": true,  "result": { "kind": "number", "value": 42 } }
{ "ok": false, "error":  { "category": "SheetNotFound", "code": 0, "message": "sheet not found: Foo" } }
```

| Field    | Always present | Description |
|----------|----------------|-------------|
| `ok`     | yes            | `true` on success, `false` on failure |
| `result` | on success     | Operation-specific output data (shape varies by op) |
| `error`  | on failure     | Structured `ExcelError` object (see §4) |

`result` and `error` are mutually exclusive.

---

## 4. Error Model

```json
{
  "category": "FileNotFound",
  "code":     2,
  "message":  "file not found: C:/data/book.xlsx",
  "hint":     "Ensure the file exists before calling range.read."
}
```

| Field      | Required | Description |
|------------|----------|-------------|
| `category` | yes      | One of the 10 PascalCase categories below |
| `code`     | yes      | Numeric code; 0 if not applicable; COM backends use the HRESULT decimal value |
| `message`  | yes      | Human-readable description |
| `hint`     | no       | Optional actionable hint for the caller |

### Error Categories

| Category             | When to use |
|----------------------|-------------|
| `FileNotFound`       | `path` does not exist on disk |
| `FileLocked`         | File is open and locked by another process |
| `InvalidArg`         | A required parameter is missing, the wrong type, or has an illegal value |
| `RangeParseError`    | `target.range` is not valid A1 notation |
| `SheetNotFound`      | `target.sheet` does not exist in the workbook |
| `UnsupportedFormat`  | `save_as.format` or the operation itself is not supported by this backend |
| `ComError`           | Unclassified COM/OLE HRESULT failure |
| `MacroTrustDisabled` | VBA trust access is required but not enabled in Excel's Trust Center |
| `Timeout`            | Operation exceeded the configured timeout |
| `Unknown`            | Any failure that does not map to the above categories |

A backend that cannot perform an operation (e.g. OpenXML cannot recalculate formulas; VBA cannot
process a file that isn't open) should return an `OpResponse` with `ok: false` and an appropriate
category — commonly `UnsupportedFormat` or `Unknown`. capctl records this as ❌ in the matrix,
which is meaningful comparison data. **Do not crash or exit non-zero** for expected limitations;
express them as structured error responses.

---

## 5. Stateless Semantics

Every mutating op follows the lifecycle: **open → modify → save → close**.

- If `save_as` is absent, infer the format from `path`'s extension (`.xlsx`→51, `.xlsm`→52,
  `.xls`→56, `.csv`→6). Fall back to xlsx (51) if the extension is unrecognized.
- If `save_as` is present, save to `save_as.path` using the given format code, then close.
  `save_as` always wins over `path` for the output location.
- `range.read` (and other read-only ops) open the file, read the value, and close without saving.
- Ops that require an existing file (`range.read`, `range.clear`, `range.copy_values`) must return
  `FileNotFound` if `path` does not exist.
- `cell.write` and `range.write_bulk` create the file if it does not exist.

**Format codes** (Excel `FileFormat` enum):

| Format string | Code |
|---------------|------|
| `xlsx`        | 51   |
| `xlsm`        | 52   |
| `xls`         | 56   |
| `csv`         | 6    |

---

## 6. Required Op Set

To produce non-⬜ entries in the capability matrix a backend must implement the following five
ops (the current Cell I/O domain). The parameters and results listed here are exact — they match
what the Rust reference backend implements in `languages/rust/src/ops/dispatch.rs`.

### `cell.write`

Write a single value to one cell.

**Params:**

| Param    | Required | Values | Description |
|----------|----------|--------|-------------|
| `kind`   | yes      | `string` \| `number` \| `bool` \| `formula` | How to interpret `value` |
| `value`  | yes      | JSON string / number / bool | The value to write |

- `kind=string` — write the value via `Range.Value2` as a string.
- `kind=number` — value must be a JSON number; write via `Range.Value2`.
- `kind=bool` — value must be a JSON boolean; write via `Range.Value2`.
- `kind=formula` — value must be a string starting with `=`; write via `Range.Formula`.

**Result (ok):** `{ "written": true }`

**Error categories:** `InvalidArg` (missing/wrong params), `ComError` (COM failure),
`UnsupportedFormat` (unknown `save_as.format`).

**File creation:** creates the workbook if it does not exist.

---

### `range.read`

Read a property of a single cell or range.

**Params:**

| Param      | Required | Values | Default |
|------------|----------|--------|---------|
| `property` | no       | `value2` \| `formula` \| `text` | `value2` |

- `property=value2` — reads `Range.Value2`; returns the computed numeric value for formula cells.
- `property=formula` — reads `Range.Formula`; returns the formula string (e.g. `=2+3`) or the
  literal value if the cell contains no formula.
- `property=text` — reads `Range.Text`; returns the displayed string as formatted by Excel.

**Result (ok):** `{ "kind": <"string"|"number"|"bool"|"empty"|"unknown">, "value": <scalar|null> }`

**Error categories:** `FileNotFound` (path absent), `SheetNotFound` (sheet not found when sheet
name was specified), `InvalidArg` (unknown `property` value), `ComError`.

**File requirement:** file must exist; returns `FileNotFound` otherwise.

---

### `range.write_bulk`

Write a rectangular 2D array of values to a range in one call.

**Params:**

| Param    | Required | Description |
|----------|----------|-------------|
| `values` | yes      | 2D JSON array (array-of-arrays). All rows must have equal length (rectangular). Cells may be string, number, bool, or `null` (null → empty cell). |

`target.range` must be a multi-cell range (must contain `:`), e.g. `"A1:B3"`.

**Result (ok):** `{ "written": true, "rows": <int>, "cols": <int> }`

**Error categories:** `InvalidArg` (missing values, non-rectangular array, single-cell target),
`ComError`.

**File creation:** creates the workbook if it does not exist.

---

### `range.clear`

Clear the contents (or all formatting+contents) of a range.

**Params:**

| Param  | Required | Values | Default |
|--------|----------|--------|---------|
| `mode` | no       | `contents` \| `all` | `contents` |

- `mode=contents` — clears cell values only (`Range.ClearContents`); preserves formatting.
- `mode=all` — clears values and formatting (`Range.Clear`).

**Result (ok):** `{ "cleared": true }`

**Error categories:** `FileNotFound` (path absent), `InvalidArg` (unknown mode), `ComError`.

**File requirement:** file must exist; returns `FileNotFound` otherwise.

---

### `range.copy_values`

Copy the `Value2` of a source cell to a destination cell (no clipboard).

**Params:**

| Param  | Required | Description |
|--------|----------|-------------|
| `dest` | yes      | A1 address of the destination cell on the same sheet |

`target.range` is the source cell address; `params.dest` is the destination.  
Both source and destination must be on the same sheet (`target.sheet`).  
The copy is value-only — formulas, formatting, and other metadata are not transferred.

**Result (ok):** `{ "copied": true }`

**Error categories:** `FileNotFound` (path absent), `InvalidArg` (missing `dest`, unsupported
source value type), `ComError`.

**File requirement:** file must exist; returns `FileNotFound` otherwise.

---

## 7. How capctl Drives a Backend

For each capability in the catalog, `capctl verify` runs the following sequence:

```
1. Delete any temp file from a previous run (fresh workbook each time).
2. Setup phase: for each op in verify.setup[], run it against the BACKEND under test.
   If any setup op fails → record ❌ (BackendFail) and skip.
3. Action phase: run verify.action against the BACKEND under test.
   If action fails → record ❌ (BackendFail) and skip.
4. Assert phase: run verify.assert.read against the REFERENCE READER (the Rust excel-ops binary).
   If the read fails → record ❌ (ReadFail).
5. Compare result.value from the reference read to verify.assert.expect:
   - Numbers: |got − expected| ≤ tol  (default tol = 1e-9)
   - Strings / bools: exact equality
   - null == null
   If match → ✅  (Pass)
   If mismatch → ⚠️  (Lossy — value was written but read back differently)
```

The reference reader is always the Rust `excel-ops` binary, so all backends are judged by one
consistent reader. An approach that cannot perform an op should return an `OpResponse` error;
capctl records that as ❌, which is still meaningful comparison data.

**Matrix symbols:**

| Symbol | Meaning |
|--------|---------|
| ✅     | Full — action succeeded and read-back matched expected |
| ⚠️     | Partial/lossy — action succeeded but read-back differed from expected |
| ❌     | Unsupported/errored — backend returned an error response, or reference reader failed |
| ⬜     | Untested — `capctl verify` has not been run for this backend |

---

## 8. Registering a New Backend

1. Build an executable that implements the protocol above.
2. Run:

```
capctl verify \
  --backend    <name>          \   # one of: vba, cpp, rust, openxml
  --backend-cmd  <path/to/exe> \   # the new backend executable
  --reference-cmd <path/to/rust-excel-ops-exe> \
  --catalog    spec/capabilities/catalog \
  --out        spec/capabilities/support-matrix.md
```

- `--backend` must be one of `vba`, `cpp`, `rust`, `openxml` (case-insensitive).
- `--reference-cmd` must point to the Rust `excel-ops` binary so that all backends are read-back
  verified by the same consistent reader.
- `--out` is overwritten with the new matrix (per-run; future work will merge multi-backend runs).

capctl creates a temporary workdir under `%TEMP%\capctl\<backend>\` for intermediate `.xlsx` files
produced during verification.

---

## Appendix: Planned Ops (not yet implemented)

The following ops appear in the design but are **not implemented** in the current Rust reference
backend and are therefore not yet verifiable:

- `range.read_bulk` — read a 2D array back from a multi-cell range (deferred; array read-back
  pending SAFEARRAY unwrap in the reference reader).
- `batch.apply` — run multiple sub-ops in a single open/close cycle (defined in contract.md;
  not yet dispatched).
- All ops in the 7 non-Cell-I/O domains (formatting, row/column structure, multi-sheet, data
  operations, charts, PDF export, VBA macros).

Do not implement these in a new backend and expect them to be verified yet.
