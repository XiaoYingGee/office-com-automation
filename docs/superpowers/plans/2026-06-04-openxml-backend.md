# OpenXML Backend Implementation Plan

> **For agentic workers:** REQUIRED: Use superpowers:subagent-driven-development (if subagents available) or superpowers:executing-plans to implement this plan. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Implement the **OpenXML** approach (#2 of 4) as a backend that satisfies `docs/backend-protocol.md` — a .NET console exe that reads one `OpRequest` JSON from stdin and writes one `OpResponse` JSON to stdout, manipulating `.xlsx` files directly via the Open XML SDK **without launching Excel** — then fill its column in the capability support matrix.

**Architecture:** A new .NET 10 console project `languages/openxml/` (`DocumentFormat.OpenXml` NuGet, `System.Text.Json`). It implements only the WRITE/setup ops (`cell.write`, `range.write_bulk`, `range.clear`, `range.copy_values`); the assert read stays the Rust `excel-ops` reference reader (Excel COM), so all approaches are judged by one consistent reader. `dotnet build` emits a native `ExcelOps.exe` apphost that `capctl --backend openxml --backend-cmd <…ExcelOps.exe>` drives.

**Tech Stack:** C# / .NET 10, DocumentFormat.OpenXml (3.x, restores as `16.0.x` package), System.Text.Json, xUnit for unit tests. Rust `capctl` + `excel-ops` (already built) for cross-approach verification.

**Key design decision (honest capability measurement):** OpenXML has **no calculation engine**. For `CELL-WRITE-FORMULA` the backend writes only the formula text (`<f>2+3</f>`) and does NOT fabricate a cached `<v>` result. Whether that capability lands ✅ or ⚠️/❌ is then determined by the Rust reference reader (Excel may recalc on open) — we record reality, we do not fake it.

---

## Reference & conventions
- Contract: `docs/backend-protocol.md` (READ THIS FIRST — it is the authoritative op spec), schemas in `spec/capabilities/schema/`.
- The Rust reference backend `languages/rust/src/ops/dispatch.rs` is the behavioral oracle — match its op semantics (params, results, error categories) exactly.
- Error categories (PascalCase strings, must match): FileNotFound, FileLocked, InvalidArg, RangeParseError, SheetNotFound, UnsupportedFormat, ComError, MacroTrustDisabled, Timeout, Unknown. (OpenXML won't produce ComError; use Unknown/UnsupportedFormat/InvalidArg/FileNotFound as appropriate.)
- New project lives at `languages/openxml/`. Build artifacts (`bin/`, `obj/`) are already git-ignored by the root `.gitignore`; SOURCE files (`*.cs`, `*.csproj`) are tracked.
- Working dir for all commands: `D:\Workspace\office-com-automation`. Branch: create `user/wuyin/openxml-backend` from `main` at the start (do NOT work on `main`).
- Commit messages end with: `Co-Authored-By: Claude Opus 4.8 <noreply@anthropic.com>`.

## File structure
```
languages/openxml/
├── ExcelOps.csproj          # console app, net10.0, DocumentFormat.OpenXml ref
├── src/
│   ├── Program.cs           # stdin→OpRequest→dispatch→OpResponse→stdout
│   ├── Contract.cs          # OpRequest/OpResponse/ExcelError records + JSON opts
│   ├── ExcelError.cs        # ErrorCategory enum + OpException
│   ├── A1.cs                # A1 address parsing (col letters↔index, range split)
│   └── Ops.cs               # the op implementations over SpreadsheetDocument
├── tests/
│   └── OpenXmlOps.Tests.csproj + *.cs   # xUnit, write-then-read-back-with-OpenXML
└── README.md                # how to build + run + its place in the matrix
```

---

## Chunk 1: Project scaffold + contract + protocol loop + `cell.write` (string/number/bool)

**Files:** create `languages/openxml/ExcelOps.csproj`, `src/Program.cs`, `src/Contract.cs`, `src/ExcelError.cs`, `src/A1.cs`, `src/Ops.cs`, `tests/OpenXmlOps.Tests.csproj`, `tests/CellWriteTests.cs`.

- [ ] **Step 1: Create branch + project skeleton**
```bash
git checkout main && git pull --ff-only
git checkout -b user/wuyin/openxml-backend
cd languages/openxml
dotnet new console -n ExcelOps -o . --force
dotnet add package DocumentFormat.OpenXml
```
Set `ExcelOps.csproj` `<AssemblyName>ExcelOps</AssemblyName>`, `<Nullable>enable</Nullable>`, `<TargetFramework>net10.0</TargetFramework>`. Confirm `dotnet build` emits `bin/Debug/net10.0/ExcelOps.exe`.

- [ ] **Step 2: Contract types (`src/Contract.cs`, `src/ExcelError.cs`)**
Define records mirroring the JSON contract, using `System.Text.Json` with snake/lower field names exactly as the protocol (`op`, `path`, `target`, `params`, `save_as`, `ok`, `result`, `error`, `category`, `code`, `message`, `hint`, `sheet`, `range`). Use `[JsonPropertyName(...)]` where C# casing differs (e.g. `SaveAs`→`save_as`). `params` and `result` are `JsonElement?` / `JsonNode?` (free-form). `ErrorCategory` is an enum serialized as its PascalCase name (e.g. `JsonStringEnumConverter`). Add an `OpException(ErrorCategory, string message, string? hint=null)` you can throw from op code and catch at the top to build an error `OpResponse`.

- [ ] **Step 3: Protocol loop (`src/Program.cs`) — write the failing test first**
Create the xUnit test project and a test that runs the built exe as a subprocess (mirrors the Rust integration tests):
```csharp
// tests/CellWriteTests.cs
static (int code, string stdout) RunOp(string json) {
    var exe = Path.GetFullPath("../../../../bin/Debug/net10.0/ExcelOps.exe"); // adjust to built path
    var psi = new ProcessStartInfo(exe){ RedirectStandardInput=true, RedirectStandardOutput=true };
    var p = Process.Start(psi)!; p.StandardInput.Write(json); p.StandardInput.Close();
    var outp = p.StandardOutput.ReadToEnd(); p.WaitForExit();
    return (p.ExitCode, outp);
}

[Fact]
public void WriteString_ReturnsOk_AndOpenXmlReadsItBack() {
    var dir = Path.Combine(Path.GetTempPath(), "oxcap"); Directory.CreateDirectory(dir);
    var path = Path.Combine(dir, "s.xlsx"); File.Delete(path);
    var p = path.Replace("\\","/");
    var (code, outp) = RunOp($@"{{""op"":""cell.write"",""path"":""{p}"",""target"":{{""sheet"":""Sheet1"",""range"":""A1""}},""params"":{{""value"":""hi"",""kind"":""string""}},""save_as"":{{""path"":""{p}"",""format"":""xlsx""}}}}");
    Assert.Contains("\"ok\":true", outp);
    // Read back WITH OpenXML (no Excel): assert A1 == "hi"
    Assert.Equal("hi", OpenXmlReadback.ReadCellString(path, "Sheet1", "A1"));
}
```
Add `OpenXmlReadback.ReadCellString` helper in the test project (opens the file with OpenXML and returns the resolved cell string, resolving shared strings). Run `dotnet test` → FAILS (Program is a stub).
> Note: the exe path resolution in tests is fiddly; prefer passing the exe path via an env var set in the csproj, or compute from `AppContext.BaseDirectory`. Get this working before moving on.

- [ ] **Step 4: Implement Program.cs + `cell.write` for string/number/bool (`src/Ops.cs`, `src/A1.cs`)**
`Program.Main`: read all of stdin, `JsonSerializer.Deserialize<OpRequest>`, `try { var result = Dispatch(req); WriteOk(result); } catch (OpException ex) { WriteErr(ex); } catch (Exception ex) { WriteErr(Unknown, ex.Message); }`. Write compact JSON to stdout.
`Ops.Dispatch(req)` switches on `req.Op`:
- `cell.write`: open-or-create the `.xlsx` at `req.Path` (if file missing, create a new SpreadsheetDocument with one sheet named per `target.sheet` or "Sheet1"); resolve/create the worksheet by `target.sheet` (fallback to first sheet); resolve the cell by `target.range` (an A1 ref); set value by `params.kind`:
  - `string` → store as a shared string (or inline string), `DataType = CellValues.String`/`SharedString`.
  - `number` → `CellValue = n.ToString(InvariantCulture)`, `DataType = CellValues.Number`.
  - `bool` → `CellValue = b ? "1":"0"`, `DataType = CellValues.Boolean`.
  - (formula handled in Chunk 2.)
  Then save to `save_as.path` (or `req.Path`) — for OpenXML "save as a different path" = save then `File.Copy`, or build at the target path directly. Honor format: only `xlsx`/`xlsm` are real OpenXML formats; `xls`/`csv` are NOT OpenXML — return `UnsupportedFormat` for those (Excel-binary/csv are out of OpenXML's scope; this is legitimate ❌ data). Return `{"written":true}`.
`A1.cs`: parse `"A1"` → (colIndex, rowIndex); column letters↔number; later a range `"A1:B2"` splitter. (DocumentFormat.OpenXml has no Range; you place `<c r="A1">` cells into `SheetData` rows in ascending row/col order — implement an `EnsureCell(worksheet, "A1")` helper that inserts a Cell at the correct position, since OpenXML requires cells in order.)
Run `dotnet test` → the WriteString test passes; add and pass WriteNumber and WriteBool tests (read back with OpenXML, asserting the typed value).

- [ ] **Step 5: Commit**
```bash
cd D:/Workspace/office-com-automation
git add languages/openxml/ExcelOps.csproj languages/openxml/src languages/openxml/tests
git commit -m "feat(openxml): project scaffold + contract loop + cell.write string/number/bool"
```

---

## Chunk 2: Remaining ops — formula, `range.write_bulk`, `range.clear`, `range.copy_values`

**Files:** modify `src/Ops.cs`, `src/A1.cs`; add tests in `tests/`.

- [ ] **Step 1: `cell.write` kind=formula (TDD)**
Test: write `params {value:"=2+3", kind:"formula"}` to A1, read back with OpenXML and assert the cell has a `<f>` element equal to `2+3` (OpenXML stores formula WITHOUT the leading `=`) and NO cached `<v>`. Implement: set `cell.CellFormula = new CellFormula("2+3")` (strip a leading `=`), do not set `CellValue`. Run → pass.
> Honest-measurement note: do NOT write a cached value. The matrix outcome for CELL-WRITE-FORMULA will be decided later by the Excel reference reader.

- [ ] **Step 2: `range.write_bulk` (TDD)**
Test: `params {values:[[10,20],[30,40]]}` to `A5:B6`; read back A6==30 and B5==20 via OpenXML. Implement: require `target.range` contain `:` (else InvalidArg); validate the 2D array is rectangular (every row same length, non-empty) → InvalidArg with the offending row index otherwise; iterate rows×cols writing each scalar (string→String, number→Number, bool→Boolean, null→empty cell) at the anchor + offset cell ref. Return `{"written":true,"rows":r,"cols":c}`. Run → pass.

- [ ] **Step 3: `range.clear` (TDD)**
Test: create a file with A1 set, then `range.clear` A1 `{mode:"contents"}`, read back A1 empty/absent via OpenXML. Implement: require existing file (FileNotFound if missing). `contents` → remove the cell's value/formula (clear `CellValue`/`CellFormula`, or remove the `<c>` inner). `all` → also remove style (`StyleIndex`). Return `{"cleared":true}`. Run → pass.

- [ ] **Step 4: `range.copy_values` (TDD)**
Test: file with H1=7, `range.copy_values` target `H1` `params {dest:"H2"}`, read back H2==7 via OpenXML. Implement: require existing file; require `params.dest`; read source cell's resolved value (handle shared strings), write same typed value to dest on the same sheet. Same-sheet only (a sheet-qualified dest is InvalidArg). Return `{"copied":true}`. Run → pass.

- [ ] **Step 5: Error-path tests + commit**
Add tests: missing file for `range.clear`/`range.copy_values` → response `ok:false` + `category:"FileNotFound"`; `range.write_bulk` ragged array → `InvalidArg`; `save_as.format:"csv"` → `UnsupportedFormat`. Run `dotnet test` (all pass).
```bash
git add languages/openxml/src languages/openxml/tests
git commit -m "feat(openxml): formula/write_bulk/clear/copy_values ops + error paths"
```

---

## Chunk 3: Cross-approach verification (capctl) + matrix + docs

**Files:** regenerate `spec/capabilities/support-matrix.md`; create `languages/openxml/README.md`; minor note in root `README.md`.

- [ ] **Step 1: Build the OpenXML exe + run capctl against it**
```bash
dotnet build languages/openxml -c Debug
# build the rust bins if needed
cargo build --manifest-path languages/rust/Cargo.toml --bins
cd languages/rust
cargo run --bin capctl -- verify \
  --backend openxml \
  --backend-cmd ../../languages/openxml/bin/Debug/net10.0/ExcelOps.exe \
  --reference-cmd target/debug/excel-ops.exe \
  --catalog ../../spec/capabilities/catalog \
  --out ../../spec/capabilities/support-matrix.md
```
This runs each Cell I/O capability through the OpenXML backend (write/setup) and the Rust excel-ops reference reader (assert), filling the **OpenXML column** while the **matrix-merge preserves the existing Rust ✅ column**. Inspect the matrix.

- [ ] **Step 2: Record + explain the outcomes (this IS the deliverable)**
For every OpenXML result that is ⚠️ or ❌, confirm it reflects a genuine OpenXML limitation (e.g. anything needing the calc engine, or `xls`/`csv` formats) and ensure capctl wrote a footnote. Do NOT "fix" a legitimate ❌ by faking data. Expected (verify, don't assume): plain value writes (string/number/bool), write_bulk, clear, copy, read-formula-text, read-text → ✅; CELL-WRITE-FORMULA → whatever Excel-on-open yields (✅ if Excel recalcs the cached-less formula, else ⚠️/❌ — record reality). If any "should-be-✅" op is ❌, investigate the OpenXML writer (likely a malformed file the Rust reader rejects) and fix the writer.

- [ ] **Step 3: Verify matrix integrity (merge didn't wipe Rust)**
`git diff spec/capabilities/support-matrix.md` — the Rust column must still be ✅ for all 9 rows; the OpenXML column is now filled; VBA/C++ stay ⬜. Re-run capctl for openxml a 2nd time → matrix is byte-stable (deterministic).

- [ ] **Step 4: Docs**
`languages/openxml/README.md`: what it is (no-Excel, file-level), how to build (`dotnet build`), how it's driven by capctl, and the honest-formula caveat. Add one line to root `README.md` capability section noting the OpenXML backend now exists and is verified for Cell I/O (link the matrix). Do NOT overclaim other domains.

- [ ] **Step 5: Commit**
```bash
git add spec/capabilities/support-matrix.md languages/openxml/README.md README.md
git commit -m "feat(openxml): verify Cell I/O via capctl + fill matrix column + docs"
```

---

## Completion criteria
- `dotnet test languages/openxml` green (writer unit tests, no Excel needed).
- `dotnet build languages/openxml` emits `ExcelOps.exe`.
- `capctl verify --backend openxml` runs clean and fills the OpenXML column; Rust column preserved (merge works); outcomes are truthful (legitimate ❌/⚠️ kept, with footnotes).
- Matrix output is deterministic across repeated runs.
- `languages/openxml/README.md` + root README note added; no overclaiming.

## Notes / cautions
- The Open XML SDK is exacting: cells must be inserted into `SheetData` rows in ascending row, then ascending column order, or Excel/readers reject the file. Implement and unit-test an `EnsureCell`/`InsertCellInOrder` helper early — most OpenXML "the file is corrupt" bugs come from out-of-order cells or a missing shared-string table.
- Strings: either a `SharedStringTable` (correct, what Excel does) or inline strings (`<is>`). Inline strings are simpler and read back fine; either is acceptable — pick one and unit-test the round-trip.
- `save_as` to a different path: simplest is to operate on a temp/the target path directly; if `req.path` and `save_as.path` differ, write/flush then `File.Copy(src, dest, overwrite:true)`.
- Do NOT add a calc engine or fake cached formula values. The whole point is to measure what OpenXML genuinely supports.
- Reference reader stays Rust `excel-ops` — do NOT implement `range.read` in OpenXML (out of scope; capctl never calls the OpenXML backend for the assert read).
- If a subagent gets stuck on an OpenXML SDK API specific, it should look it up / write a tiny spike, not guess — corrupt .xlsx files fail opaquely.

## Next plans (not in this one)
- `…-cpp-backend.md` (approach #3, COM via vcvarsall+cl), `…-vba-backend.md` (approach #4, script host).
- `…-domain-rollout-formatting.md` etc.: once ≥2 backends exist, roll new capabilities per-domain, each implemented across all available backends at once (per the user's "每条目立即多方案" directive).
