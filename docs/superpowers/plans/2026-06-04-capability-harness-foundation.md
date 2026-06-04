# Capability Harness Foundation 实现计划

> **For agentic workers:** REQUIRED: Use superpowers:subagent-driven-development (if subagents available) or superpowers:executing-plans to implement this plan. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 立起"语言无关能力契约 + 编排器 + 参考读取器 + Rust backend"的端到端脊柱，跑通"一条能力 → Rust 实现 → 参考读取器自动验证 → 回填支持矩阵一行"的闭环。

**Architecture:** 契约用**进程边界**落地——每个方案是一个可执行程序，从 stdin 读 `OpRequest` JSON、往 stdout 写 `OpResponse` JSON。编排器 `capctl`（Rust）按"能力 × backend"驱动：跑 backend 的 write `action`，再用**参考读取器**（Rust `excel-ops read`，复用已验证的 `com.rs`）重开断言，判 ✅/⚠️/❌ 并回填 `support-matrix.md`。本计划只立脊柱 + Rust backend；OpenXML/C++/VBA backend 为后续各自一份计划。

**Tech Stack:** Rust（`windows` crate + 复用现有 `languages/rust/src/com.rs`、`excel.rs`）、JSON（`serde_json`）、JSON Schema（校验能力条目）。

**前置说明（排序）:** "每条新能力立即 4 方案实现"是**条目滚动阶段**（域计划，本计划之后）的规则。本计划与紧随的 OpenXML/C++/VBA backend 计划属一次性**基础设施批**，把 4 个 backend 立到同一套 JSON 契约后，再开始 4 方案逐条目滚动。

---

## 参考与约定

- 设计稿：`docs/superpowers/specs/2026-06-04-excel-capability-catalog-design.md`
- 复用：`languages/rust/src/com.rs`（`Dispatch`/`Variant`/`ComGuard`，已验证）、`excel.rs`（`ExcelApp`）
- 新增 Rust 组件落在现有 `languages/rust` 工程内（同一 `Cargo.toml`），新增：
  - 库模块 `src/ops/`（契约类型 + dispatcher）
  - bin `src/bin/excel-ops.rs`（backend：stdin JSON → stdout JSON，含 write 与 read 两类 op）
  - bin `src/bin/capctl.rs`（编排器：跑能力 × backend，参考读取器断言，回填矩阵）
- 契约 schema 与能力表落在 `spec/capabilities/`
- 提交粒度：每个 Step 末尾的 commit 即时提交；分支沿用当前 `user/wuyin/rust-excel-com-tasks`

---

## Chunk 1: 契约类型 + schema + capctl 骨架

**Files:**
- Create: `spec/capabilities/schema/op-request.schema.json`
- Create: `spec/capabilities/schema/error.schema.json`
- Create: `spec/capabilities/schema/capability.schema.json`
- Create: `spec/capabilities/contract.md`
- Modify: `languages/rust/Cargo.toml`（加 `serde`/`serde_json` 依赖、声明 bin）
- Create: `languages/rust/src/ops/mod.rs`（契约类型：`OpRequest`/`OpResponse`/`ExcelError`/`ErrorCategory`）
- Test: `languages/rust/src/ops/mod.rs`（`#[cfg(test)]` 往返序列化）

- [ ] **Step 1: 写契约类型的失败测试**

在 `languages/rust/src/ops/mod.rs` 写：
```rust
use serde::{Deserialize, Serialize};

#[derive(Debug, Serialize, Deserialize, PartialEq)]
pub struct OpRequest {
    pub op: String,
    pub path: String,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub target: Option<Target>,
    #[serde(default)]
    pub params: serde_json::Value,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub save_as: Option<SaveAs>,
}

#[derive(Debug, Serialize, Deserialize, PartialEq)]
pub struct Target {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sheet: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub range: Option<String>,
}

#[derive(Debug, Serialize, Deserialize, PartialEq)]
pub struct SaveAs { pub path: String, pub format: String }

#[derive(Debug, Serialize, Deserialize, PartialEq)]
#[serde(tag = "ok")]
pub enum OpResponse {
    #[serde(rename = "true")]
    Ok { result: serde_json::Value },
    #[serde(rename = "false")]
    Err { error: ExcelError },
}

#[derive(Debug, Serialize, Deserialize, PartialEq)]
pub struct ExcelError {
    pub category: ErrorCategory,
    pub code: i64,
    pub message: String,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub hint: Option<String>,
}

#[derive(Debug, Serialize, Deserialize, PartialEq)]
#[serde(rename_all = "PascalCase")]
pub enum ErrorCategory {
    FileNotFound, FileLocked, InvalidArg, RangeParseError, SheetNotFound,
    UnsupportedFormat, ComError, MacroTrustDisabled, Timeout, Unknown,
}

#[cfg(test)]
mod tests {
    use super::*;
    #[test]
    fn op_request_roundtrips() {
        let json = r#"{"op":"cell.write","path":"x.xlsx","target":{"sheet":"Sheet1","range":"A1"},"params":{"value":"hi","kind":"string"}}"#;
        let req: OpRequest = serde_json::from_str(json).unwrap();
        assert_eq!(req.op, "cell.write");
        let back = serde_json::to_string(&req).unwrap();
        let req2: OpRequest = serde_json::from_str(&back).unwrap();
        assert_eq!(req, req2);
    }
}
```

> **注意 `#[serde(tag = "ok")]` 的布尔标签**：serde 的内部 tag 需字符串，契约里 `ok` 是布尔。实现时若 `tag="ok"` 无法匹配布尔，改为自定义：`OpResponse` 用结构体 `{ ok: bool, result?, error? }` + 手写校验，或在 Step 3 调整。先按上方写测试，Step 2 跑红后据实际报错在 Step 3 决定最终形态（这是真实的设计校验点，不要跳过）。

- [ ] **Step 2: 加依赖并跑测试看失败**

`languages/rust/Cargo.toml` 增加：
```toml
[dependencies]
serde = { version = "1", features = ["derive"] }
serde_json = "1"

[lib]
path = "src/lib.rs"

[[bin]]
name = "excel-com"
path = "src/main.rs"

[[bin]]
name = "excel-ops"
path = "src/bin/excel-ops.rs"

[[bin]]
name = "capctl"
path = "src/bin/capctl.rs"
```
创建 `languages/rust/src/lib.rs`：
```rust
pub mod com;
pub mod excel;
pub mod ops;
```
> 现有 `main.rs` 用 `mod com;` 等内联模块；改为 lib 后 `main.rs`/`tasks.rs` 改用 `use excel_com::{com, excel};`。本步同时创建占位 `src/bin/excel-ops.rs` 与 `src/bin/capctl.rs`（内容 `fn main(){}`）以便 `Cargo.toml` 不报错。

Run: `cd languages/rust && cargo test ops:: 2>&1 | tail -20`
Expected: 编译/测试失败（类型或 `tag` 问题），据报错进入 Step 3。

- [ ] **Step 3: 修正契约类型直到测试通过**

据 Step 2 报错调整 `OpResponse` 表示（大概率改为）：
```rust
#[derive(Debug, Serialize, Deserialize, PartialEq)]
pub struct OpResponse {
    pub ok: bool,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub result: Option<serde_json::Value>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub error: Option<ExcelError>,
}
impl OpResponse {
    pub fn ok(result: serde_json::Value) -> Self { Self { ok: true, result: Some(result), error: None } }
    pub fn err(error: ExcelError) -> Self { Self { ok: false, result: None, error: Some(error) } }
}
```
Run: `cargo test ops:: 2>&1 | tail -20`
Expected: PASS

- [ ] **Step 4: 写 JSON Schema（契约的机器可读形态）**

`spec/capabilities/schema/op-request.schema.json`、`error.schema.json`、`capability.schema.json`，与 Rust 类型对齐。`capability.schema.json` 描述单条能力字段（`id/name/desc/op/param_path/sample/verify/errors/com_ref/support`，`support` 含 `vba/cpp/rust/openxml` 枚举 `✅|⚠️|❌|⬜`）。`contract.md` 用散文 + 上述 schema 链接说明无状态语义、错误模型、batch。

- [ ] **Step 5: Commit**
```bash
git add spec/capabilities/ languages/rust/Cargo.toml languages/rust/src/lib.rs languages/rust/src/ops/ languages/rust/src/bin/
git commit -m "feat(cap): contract types + JSON schemas + harness scaffolding"
```

---

## Chunk 2: Rust backend `excel-ops`（write + 参考读取器 read）

**Files:**
- Create/replace: `languages/rust/src/bin/excel-ops.rs`
- Create: `languages/rust/src/ops/dispatch.rs`（`execute(OpRequest, &ExcelApp) -> OpResponse`）
- Modify: `languages/rust/src/ops/mod.rs`（`pub mod dispatch;`）
- Test: `languages/rust/tests/excel_ops_cli.rs`（集成测试：spawn bin，喂 JSON）

- [ ] **Step 1: 写 backend CLI 的失败集成测试**

`languages/rust/tests/excel_ops_cli.rs`：
```rust
use std::process::{Command, Stdio};
use std::io::Write;

fn run_op(json: &str) -> String {
    let mut child = Command::new(env!("CARGO_BIN_EXE_excel-ops"))
        .stdin(Stdio::piped()).stdout(Stdio::piped()).spawn().unwrap();
    child.stdin.take().unwrap().write_all(json.as_bytes()).unwrap();
    let out = child.wait_with_output().unwrap();
    String::from_utf8(out.stdout).unwrap()
}

#[test]
fn write_then_reference_read_roundtrip() {
    let dir = std::env::temp_dir().join("capfound");
    std::fs::create_dir_all(&dir).unwrap();
    let path = dir.join("rt.xlsx");
    let p = path.to_string_lossy().replace('\\', "/");
    let w = format!(r#"{{"op":"cell.write","path":"{p}","target":{{"sheet":"Sheet1","range":"A1"}},"params":{{"value":"hi","kind":"string"}},"save_as":{{"path":"{p}","format":"xlsx"}}}}"#);
    let wr = run_op(&w);
    assert!(wr.contains("\"ok\":true"), "write resp: {wr}");
    let r = format!(r#"{{"op":"range.read","path":"{p}","target":{{"sheet":"Sheet1","range":"A1"}},"params":{{}}}}"#);
    let rr = run_op(&r);
    assert!(rr.contains("hi"), "read resp: {rr}");
}
```

- [ ] **Step 2: 跑测试看失败**
Run: `cd languages/rust && cargo test --test excel_ops_cli 2>&1 | tail -25`
Expected: FAIL（`excel-ops` 仍是空 `main`）

- [ ] **Step 3: 实现 dispatch + backend main（最小：`cell.write` / `range.read`）**

`src/ops/dispatch.rs`：实现 `execute`，对 `op`：
- `cell.write`：`ExcelApp::open_or_create(path)` → 取 sheet/range → 据 `params.kind`(`string|number|bool|formula`) put `Value2`/`Formula` → `save_as`（按 format 映射 51/52/csv...）→ `OpResponse::ok`。
- `range.read`：打开 → 读 `Value2` → 归一化成 `{kind,value}` JSON → ok。
- COM 错误经 `windows::core::Error` → `map_com_error()` 映射到 `ErrorCategory`（HRESULT→category 表；文件不存在→FileNotFound 等），返回 `OpResponse::err`。

`src/bin/excel-ops.rs`：
```rust
use std::io::{Read, Write};
use excel_com::com::ComGuard;
use excel_com::excel::ExcelApp;
use excel_com::ops::{OpRequest, OpResponse, dispatch};

fn main() {
    let mut buf = String::new();
    std::io::stdin().read_to_string(&mut buf).unwrap();
    let resp = match serde_json::from_str::<OpRequest>(&buf) {
        Ok(req) => {
            let _com = ComGuard::new().unwrap();
            dispatch::execute(req)
        }
        Err(e) => OpResponse::err(excel_com::ops::ExcelError{
            category: excel_com::ops::ErrorCategory::InvalidArg, code:0,
            message: format!("bad request json: {e}"), hint: None }),
    };
    let s = serde_json::to_string(&resp).unwrap();
    std::io::stdout().write_all(s.as_bytes()).unwrap();
}
```
> 需要在 `ExcelApp` 上补 `open_or_create(path)`（文件存在则 open，否则 add_workbook）。复用现有 `excel.rs` 方法；新增小函数。

- [ ] **Step 4: 跑测试看通过**
Run: `cargo test --test excel_ops_cli 2>&1 | tail -25`
Expected: PASS（写入并用同一 backend 读回 "hi"）

- [ ] **Step 5: Commit**
```bash
git add languages/rust/src/ops/ languages/rust/src/bin/excel-ops.rs languages/rust/src/excel.rs languages/rust/tests/
git commit -m "feat(cap): rust excel-ops backend with cell.write + range.read (reference reader)"
```

---

## Chunk 3: capctl 编排器 + 参考读取器断言 + 支持矩阵回填

**Files:**
- Create/replace: `languages/rust/src/bin/capctl.rs`
- Create: `languages/rust/src/ops/catalog.rs`（能力条目结构 + 加载）
- Create: `spec/capabilities/catalog/01-cell-range-io.md`（先放 2-3 条带 frontmatter/JSON 的能力）
- Create: `spec/capabilities/support-matrix.md`（capctl 生成）
- Test: `languages/rust/tests/capctl_verify.rs`

- [ ] **Step 1: 写 capctl 验证闭环的失败测试**

`capctl_verify.rs`：调用 `capctl verify --backend rust --catalog <dir> --out <matrix>`，断言退出码 0 且矩阵文件含 `rust` 列为 ✅ 的行。（能力定义见 Step 3 catalog。）

- [ ] **Step 2: 跑测试看失败**
Run: `cargo test --test capctl_verify 2>&1 | tail -25`　Expected: FAIL

- [ ] **Step 3: 写头几条能力 + 实现 capctl**

`spec/capabilities/catalog/01-cell-range-io.md`：用统一结构写 `CELL-WRITE-STRING`、`CELL-WRITE-NUMBER`、`CELL-WRITE-BOOL`，每条含 `verify`（action=cell.write、assert.read=range.read、expect、tol?）。catalog 解析：能力以 JSON code-block 或 frontmatter 嵌在 md（`catalog.rs` 解析该结构；schema 校验复用 Chunk 1 schema）。

`capctl.rs` 逻辑：
1. 加载 catalog 全部能力。
2. 对每条 × 指定 backend：
   - 组 `action` 的 OpRequest（注入临时 out 路径）→ spawn backend 子进程 → 收 OpResponse。
   - 若 `verify.reopen` → 用**参考读取器**（`excel-ops range.read`，固定 Rust）跑 `assert.read` → 拿实际值。
   - 与 `expect` 比对（数值带 `tol`）→ 判 ✅/❌（写入成功但读回不符/有损 → ⚠️）。
   - backend 直接返回 `Unsupported*`/`UnsupportedFormat` → ❌。
3. 汇总写 `support-matrix.md`（行=能力 id，列=4 方案；本次只填指定 backend 列，其余 ⬜）。

- [ ] **Step 4: 跑测试看通过**
Run: `cargo test --test capctl_verify 2>&1 | tail -25`　Expected: PASS（Rust 列 ✅）
再手动：`cargo run --bin capctl -- verify --backend rust --catalog ../../spec/capabilities/catalog --out ../../spec/capabilities/support-matrix.md` 并查看矩阵。

- [ ] **Step 5: Commit**
```bash
git add languages/rust/src/ops/catalog.rs languages/rust/src/bin/capctl.rs spec/capabilities/catalog/ spec/capabilities/support-matrix.md languages/rust/tests/
git commit -m "feat(cap): capctl orchestrator + reference-reader verify + matrix output"
```

---

## Chunk 4: Cell I/O 域补齐 + 文档衔接 + backend 接口固化

**Files:**
- Modify: `spec/capabilities/catalog/01-cell-range-io.md`（补到 ~15 条：formula、value2/text、bulk 区域读写、clear 等）
- Modify: `languages/rust/src/ops/dispatch.rs`（实现这些 op）
- Create: `spec/capabilities/README.md`、`docs/backend-protocol.md`（backend 必须实现的 JSON 契约，供 OpenXML/C++/VBA 后续计划照做）
- Modify: `README.md`、`spec/capability-matrix.md`（措辞从"COM 比武"→"Excel 自动化方案比武"；加指针指向 support-matrix）

- [ ] **Step 1**: 为每条新增能力先写/扩 `capctl` 期望（catalog 即测试口径），跑 `capctl` 看新行红。
- [ ] **Step 2**: 在 `dispatch.rs` 逐条实现对应 op，跑 `capctl` 看转 ✅。
- [ ] **Step 3**: 写 `docs/backend-protocol.md`：明确"一个 backend = 读 stdin OpRequest / 写 stdout OpResponse 的可执行文件；必须实现的 op 列表与语义；错误 category 约定"。这是 OpenXML/C++/VBA 三份后续计划的对接面。
- [ ] **Step 4**: 更新 `README.md` 与 `spec/capability-matrix.md` 措辞与指针。
- [ ] **Step 5: Commit**
```bash
git add spec/ docs/ languages/rust/ README.md
git commit -m "feat(cap): complete cell-range-io domain (rust) + backend protocol doc + repo reframe"
```

---

## 完成口径（本计划）

- `cargo test` 全绿（契约往返、backend CLI、capctl 验证）。
- `capctl verify --backend rust` 能跑通 Cell I/O 全域并生成 `support-matrix.md`（Rust 列填满 ✅/⚠️/❌，其余方案 ⬜）。
- `docs/backend-protocol.md` 固化 backend 对接面。
- 脊柱可复用：后续 OpenXML/C++/VBA backend 只需各自产出一个"读 stdin OpRequest / 写 stdout OpResponse"的可执行文件，即可被 `capctl --backend <x>` 验证并回填矩阵列。

## 后续计划（各自一份，不在本计划内）

1. `2026-..-openxml-backend.md`：.NET Open XML SDK backend（`languages/openxml/`），实现 backend 协议；`capctl --backend openxml` 回填列（预期公式/图表/PDF 等域大面积 ❌）。
2. `2026-..-cpp-backend.md`：C++ backend（复用 `vcvarsall+cl`）。
3. `2026-..-vba-backend.md`：VBA backend（脚本宿主 runner）。
4. `2026-..-domain-rollout-*.md`：4 backend 齐备后，按域逐条目纵切，每条 4 方案一起实现 + 验证 + 回填。
