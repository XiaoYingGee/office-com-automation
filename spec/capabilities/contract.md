# Capability Contract

This document describes the stateless JSON contract used by the Excel capability harness.
See also: the three JSON Schema files in `spec/capabilities/schema/`.

## Semantics

Every operation follows the same open → modify → save → close lifecycle, even if the caller
does not explicitly request a save. Operations are **stateless**: each `OpRequest` receives a
fully-resolved `path` and the implementation opens, operates, and closes the workbook within a
single call. No session state is maintained between calls.

Callers that need multiple mutations in a single open/close cycle use the `batch.apply` op,
which accepts an array of sub-operations in `params.ops`.

## OpRequest

```json
{
  "op":   "cell.write",
  "path": "C:/data/book.xlsx",
  "target": { "sheet": "Sheet1", "range": "B2" },
  "params": { "value": 42, "kind": "number" },
  "save_as": { "path": "C:/data/book-out.xlsx", "format": "xlsx" }
}
```

Fields:

| Field      | Required | Description |
|------------|----------|-------------|
| `op`       | yes      | Dotted operation name (`cell.write`, `range.read`, `batch.apply`, …) |
| `path`     | yes      | Source workbook path (absolute or resolvable from CWD) |
| `target`   | no       | `{ sheet?, range? }` — worksheet name and/or A1 range address |
| `params`   | no       | Operation-specific parameters; defaults to `{}` |
| `save_as`  | no       | If present, save output to this path/format before closing |

Schema: [`op-request.schema.json`](schema/op-request.schema.json)

## OpResponse

```json
{ "ok": true,  "result": { "value": 42 } }
{ "ok": false, "error":  { "category": "SheetNotFound", "code": 2, "message": "…" } }
```

The `ok` field is always present. On success, `result` holds operation-specific output data.
On failure, `error` holds a structured `ExcelError` object (see below). The two fields are
mutually exclusive.

## Error Model

```json
{
  "category": "FileNotFound",
  "code": 2,
  "message": "No file at C:/data/book.xlsx",
  "hint": "Check the path and ensure the file exists before calling."
}
```

Ten error categories cover the full failure space:

| Category             | When raised |
|----------------------|-------------|
| `FileNotFound`       | `path` does not exist |
| `FileLocked`         | File is open by another process |
| `InvalidArg`         | A required param is missing or wrong type |
| `RangeParseError`    | `target.range` is not valid A1 notation |
| `SheetNotFound`      | `target.sheet` does not exist in the workbook |
| `UnsupportedFormat`  | `save_as.format` is not supported by the backend |
| `ComError`           | Unclassified COM/OLE HRESULT failure |
| `MacroTrustDisabled` | VBA trust access required but not enabled |
| `Timeout`            | Operation exceeded the configured timeout |
| `Unknown`            | Any error not matching the above |

Schema: [`error.schema.json`](schema/error.schema.json)

## `batch.apply`

The `batch.apply` op runs a sequence of sub-operations inside a single open/close cycle:

```json
{
  "op":   "batch.apply",
  "path": "C:/data/book.xlsx",
  "params": {
    "ops": [
      { "op": "cell.write", "target": { "sheet": "S1", "range": "A1" }, "params": { "value": "hi" } },
      { "op": "cell.write", "target": { "sheet": "S1", "range": "A2" }, "params": { "value": 99 } }
    ]
  },
  "save_as": { "path": "C:/data/book.xlsx", "format": "xlsx" }
}
```

The response `result` is an array of per-op results in the same order. The batch fails fast on
the first sub-operation error unless `params.continue_on_error` is `true`.

## Capability Catalog

Individual capabilities are described in JSON files that conform to
[`capability.schema.json`](schema/capability.schema.json). Each entry includes a `verify`
recipe (write action + optional reopen + read-back assertion) that `capctl verify` executes
automatically to produce the cross-language support matrix.
