# Cell & Range I/O Capabilities

This file catalogs the basic cell read/write capabilities for the Excel COM automation harness.

## CELL-WRITE-STRING

Write a UTF-8 string value to a single cell via `Range.Value2`.

```json
{
  "id": "CELL-WRITE-STRING",
  "name": "Write string to cell",
  "desc": "Write a text value to a single cell via Value2 and read it back.",
  "op": "cell.write",
  "param_path": "params.value (kind=string)",
  "sample": { "target": {"sheet":"Sheet1","range":"A1"}, "params": {"value":"hi","kind":"string"} },
  "verify": {
    "action": { "op":"cell.write", "target": {"sheet":"Sheet1","range":"A1"}, "params": {"value":"Capability Probe","kind":"string"} },
    "reopen": true,
    "assert": { "read": {"op":"range.read","target":{"sheet":"Sheet1","range":"A1"}}, "expect": "Capability Probe" }
  },
  "errors": [],
  "com_ref": "Range.Value2",
  "support": {"vba":"⬜","cpp":"⬜","rust":"⬜","openxml":"⬜"}
}
```

## CELL-WRITE-NUMBER

Write a floating-point numeric value to a single cell via `Range.Value2`.

```json
{
  "id": "CELL-WRITE-NUMBER",
  "name": "Write number to cell",
  "desc": "Write a floating-point number to a single cell via Value2 and read it back.",
  "op": "cell.write",
  "param_path": "params.value (kind=number)",
  "sample": { "target": {"sheet":"Sheet1","range":"B2"}, "params": {"value":1234.5,"kind":"number"} },
  "verify": {
    "action": { "op":"cell.write", "target": {"sheet":"Sheet1","range":"B2"}, "params": {"value":1234.5,"kind":"number"} },
    "reopen": true,
    "assert": { "read": {"op":"range.read","target":{"sheet":"Sheet1","range":"B2"}}, "expect": 1234.5, "tol": 0.001 }
  },
  "errors": [],
  "com_ref": "Range.Value2",
  "support": {"vba":"⬜","cpp":"⬜","rust":"⬜","openxml":"⬜"}
}
```

## CELL-WRITE-BOOL

Write a boolean value to a single cell via `Range.Value2`.

```json
{
  "id": "CELL-WRITE-BOOL",
  "name": "Write bool to cell",
  "desc": "Write a boolean value to a single cell via Value2 and read it back.",
  "op": "cell.write",
  "param_path": "params.value (kind=bool)",
  "sample": { "target": {"sheet":"Sheet1","range":"C3"}, "params": {"value":true,"kind":"bool"} },
  "verify": {
    "action": { "op":"cell.write", "target": {"sheet":"Sheet1","range":"C3"}, "params": {"value":true,"kind":"bool"} },
    "reopen": true,
    "assert": { "read": {"op":"range.read","target":{"sheet":"Sheet1","range":"C3"}}, "expect": true }
  },
  "errors": [],
  "com_ref": "Range.Value2",
  "support": {"vba":"⬜","cpp":"⬜","rust":"⬜","openxml":"⬜"}
}
```

## CELL-WRITE-FORMULA

Write a formula to a cell and verify it evaluates to the expected numeric result.

```json
{
  "id": "CELL-WRITE-FORMULA",
  "name": "Write formula to cell",
  "desc": "Write a formula string via Range.Formula; read back Value2 to verify the computed result.",
  "op": "cell.write",
  "param_path": "params.value (kind=formula)",
  "sample": { "target": {"sheet":"Sheet1","range":"D1"}, "params": {"value":"=2+3","kind":"formula"} },
  "verify": {
    "action": { "op":"cell.write", "target": {"sheet":"Sheet1","range":"D1"}, "params": {"value":"=2+3","kind":"formula"} },
    "reopen": true,
    "assert": { "read": {"op":"range.read","target":{"sheet":"Sheet1","range":"D1"}}, "expect": 5, "tol": 0.001 }
  },
  "errors": [],
  "com_ref": "Range.Formula",
  "support": {"vba":"⬜","cpp":"⬜","rust":"⬜","openxml":"⬜"}
}
```

## CELL-READ-FORMULA

Read the formula string stored in a cell via `Range.Formula`.

```json
{
  "id": "CELL-READ-FORMULA",
  "name": "Read formula from cell",
  "desc": "After writing a formula, read it back as a formula string via params.property=formula.",
  "op": "range.read",
  "param_path": "params.property=formula",
  "sample": { "target": {"sheet":"Sheet1","range":"D2"}, "params": {"property":"formula"} },
  "verify": {
    "action": { "op":"cell.write", "target": {"sheet":"Sheet1","range":"D2"}, "params": {"value":"=2+3","kind":"formula"} },
    "reopen": true,
    "assert": { "read": {"op":"range.read","target":{"sheet":"Sheet1","range":"D2"},"params":{"property":"formula"}}, "expect": "=2+3" }
  },
  "errors": [],
  "com_ref": "Range.Formula",
  "support": {"vba":"⬜","cpp":"⬜","rust":"⬜","openxml":"⬜"}
}
```

## CELL-READ-TEXT

Read the displayed text of a cell via `Range.Text`.

```json
{
  "id": "CELL-READ-TEXT",
  "name": "Read displayed text from cell",
  "desc": "After writing a number, read the displayed string via params.property=text. May vary by locale.",
  "op": "range.read",
  "param_path": "params.property=text",
  "sample": { "target": {"sheet":"Sheet1","range":"D3"}, "params": {"property":"text"} },
  "verify": {
    "action": { "op":"cell.write", "target": {"sheet":"Sheet1","range":"D3"}, "params": {"value":1234.5,"kind":"number"} },
    "reopen": true,
    "assert": { "read": {"op":"range.read","target":{"sheet":"Sheet1","range":"D3"},"params":{"property":"text"}}, "expect": "1234.5" }
  },
  "errors": [],
  "com_ref": "Range.Text",
  "support": {"vba":"⬜","cpp":"⬜","rust":"⬜","openxml":"⬜"}
}
```

## RANGE-WRITE-BULK

Write a 2D array of values to a range in a single COM call via SAFEARRAY.

```json
{
  "id": "RANGE-WRITE-BULK",
  "name": "Write 2D array to range (bulk)",
  "desc": "Write a 2D array of scalars to a range via Range.Value2 SAFEARRAY assignment.",
  "op": "range.write_bulk",
  "param_path": "params.values (2D array)",
  "sample": { "target": {"sheet":"Sheet1","range":"A5:B6"}, "params": {"values":[[10,20],[30,40]]} },
  "verify": {
    "action": { "op":"range.write_bulk", "target": {"sheet":"Sheet1","range":"A5:B6"}, "params": {"values":[[10,20],[30,40]]} },
    "reopen": true,
    "assert": { "read": {"op":"range.read","target":{"sheet":"Sheet1","range":"A6"}}, "expect": 30, "tol": 0.001 }
  },
  "errors": [],
  "com_ref": "Range.Value2 (SAFEARRAY)",
  "support": {"vba":"⬜","cpp":"⬜","rust":"⬜","openxml":"⬜"}
}
```

## RANGE-CLEAR-CONTENTS

Clear cell contents from a range via `Range.ClearContents`.

Note: capctl runs ONE action then ONE read. Since the action is `range.clear` on a fresh (empty) cell,
the assert reads back null (empty). This verifies the op does not error and leaves the cell empty.
A more meaningful test (write then clear) requires a `setup` step — deferred to a future capctl enhancement.

```json
{
  "id": "RANGE-CLEAR-CONTENTS",
  "name": "Clear cell contents",
  "desc": "Clear the contents of a cell range via Range.ClearContents. Asserts the cell is empty after clear.",
  "op": "range.clear",
  "param_path": "params.mode=contents",
  "sample": { "target": {"sheet":"Sheet1","range":"F1"}, "params": {"mode":"contents"} },
  "verify": {
    "action": { "op":"range.clear", "target": {"sheet":"Sheet1","range":"F1"}, "params": {"mode":"contents"} },
    "reopen": true,
    "assert": { "read": {"op":"range.read","target":{"sheet":"Sheet1","range":"F1"}}, "expect": null }
  },
  "errors": [],
  "com_ref": "Range.ClearContents",
  "support": {"vba":"⬜","cpp":"⬜","rust":"⬜","openxml":"⬜"}
}
```

## RANGE-COPY-VALUES

Copy the value from one cell to another via a read+write cycle (no clipboard).

Note: capctl runs ONE action then ONE read. With a fresh file H1 is empty, so copy_values copies empty→H2,
and the read of H2 returns null/empty. This verifies the op path without error.
A more meaningful test (write to source first) requires a `setup` step — deferred to a future capctl enhancement.

```json
{
  "id": "RANGE-COPY-VALUES",
  "name": "Copy cell value to another cell",
  "desc": "Copy a cell's Value2 to a destination cell without using the clipboard.",
  "op": "range.copy_values",
  "param_path": "params.dest (destination address)",
  "sample": { "target": {"sheet":"Sheet1","range":"H1"}, "params": {"dest":"H2"} },
  "verify": {
    "action": { "op":"range.copy_values", "target": {"sheet":"Sheet1","range":"H1"}, "params": {"dest":"H2"} },
    "reopen": true,
    "assert": { "read": {"op":"range.read","target":{"sheet":"Sheet1","range":"H2"}}, "expect": null }
  },
  "errors": [],
  "com_ref": "Range.Value2",
  "support": {"vba":"⬜","cpp":"⬜","rust":"⬜","openxml":"⬜"}
}
```

