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
