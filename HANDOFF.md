# Handoff — Excel 自动化"多方案比武场" / Capability Harness

> **Date:** 2026-06-04
> **Repo:** `D:\Workspace\office-com-automation`  (GitHub: `XiaoYingGee/office-com-automation`)
> **`main` HEAD at handoff:** `daa13a7`(已推送)
> **接手方式:** 读本文件即可接续。无需读旧 handoff（已被本文件取代）。

---

## 0. 一句话现状

这个仓库用 **4 种方案**（VBA / C++ / Rust 走 COM，OpenXML 走 .NET 文件级）实现同一组 Excel 能力，横向比"谁支持得更全/更好"。本次会话完成了 **(a) Rust 的 E01–E12 旧任务集** 和 **(b) 一套语言无关的"能力表 + 契约 + 编排器"基础设施 + Rust 参考后端 + OpenXML 后端**。下一步是 **C++ 后端、VBA 后端，然后按域滚动扩充能力**。

---

## 1. 本次已完成(全部已 merge 到 `main` 并推送)

1. **Rust E01–E12**(`languages/rust/`,bin `excel-com`):旧的 12 个 Excel COM 任务,12/12 PASS。(E11 在"信任访问 VBA 工程对象模型"关闭时按 SKIP 计 PASS,与 C++ 行为一致。)
2. **Capability Harness 基础设施**(语言无关):
   - **契约**:JSON `OpRequest`/`OpResponse` over stdin/stdout + 结构化错误 `{category,code,message,hint}`。
   - **Rust 参考后端**(bin `excel-ops`):实现写类 op + `range.read`(兼**统一参考读取器**)。
   - **编排器**(bin `capctl`):跑 `setup → action → assert-read`,用参考读取器断言,产出**支持矩阵**;多后端**合并**不互相覆盖。
   - **能力表**:`spec/capabilities/catalog/01-cell-range-io.md`(Cell I/O 域 9 条)。
3. **OpenXML 后端**(`languages/openxml/`,exe `ExcelOps.exe`):.NET 10 + DocumentFormat.OpenXml,直接改 `.xlsx` 不起 Excel。

**当前支持矩阵**(`spec/capabilities/support-matrix.md`):

| 域 | VBA | C++ | Rust | OpenXML |
|---|:---:|:---:|:---:|:---:|
| Cell I/O ×9 | ⬜ | ⬜ | ✅×9 | ✅×9 |

---

## 2. 架构(接手必读)

**核心抽象:一个"后端"= 一个可执行文件**,从 stdin 读 1 个 `OpRequest` JSON,往 stdout 写 1 个 `OpResponse` JSON。无状态:每个 op = 开→改→存→关。

- **契约权威文档:`docs/backend-protocol.md`**(新后端作者照它实现)。机器可读 schema 在 `spec/capabilities/schema/`。
- **统一参考读取器**:断言阶段的"重开读回"永远由 **Rust `excel-ops`(Excel COM)** 完成——无论文件哪个后端写的,用同一把尺量。所以**新后端只需实现"写类"op**(`cell.write` / `range.write_bulk` / `range.clear` / `range.copy_values`,含 setup 用的 `cell.write`),**不需要实现 `range.read`**。
- **编排器 `capctl`**:`capctl verify --backend <名> --backend-cmd <后端exe> --reference-cmd <rust excel-ops exe> --catalog <目录> --out <matrix.md>`。它逐条能力跑 setup→action(经后端)→assert read(经参考读取器),比对 `result.value` 与 `expect`(数值带 tol;字符串/布尔精确;null==null),记 ✅/⚠️/❌,**合并**进矩阵(保留其它后端列)。
- **能力表条目结构**:`id / name / desc / op / param_path / sample / verify{setup[], action, reopen, assert{read, expect, tol}} / errors / com_ref / support{vba,cpp,rust,openxml}`。一条能力 = 一个细粒度单属性,目标全表 ~120 条,现仅 Cell I/O 9 条。
- **能力做不到的方案**应在 OpResponse 里返回错误(category 如 `UnsupportedFormat`/`Unknown`),capctl 记 ❌——这是有价值的对比数据(预期 OpenXML 在公式重算/图表/PDF/宏域大面积 ❌)。

**设计稿全文**:`docs/superpowers/specs/2026-06-04-excel-capability-catalog-design.md`(含全部决策:意图原子操作/无状态/结构化错误/细粒度/自动验证/4 方案/批的单位=能力条目)。

---

## 3. 关键文件地图

```
docs/
  backend-protocol.md                        # ★新后端实现契约
  superpowers/specs/2026-06-04-excel-capability-catalog-design.md   # ★设计稿
  superpowers/plans/2026-06-04-capability-harness-foundation.md     # 已执行
  superpowers/plans/2026-06-04-openxml-backend.md                   # 已执行
spec/capabilities/
  contract.md  README.md  support-matrix.md  # ★矩阵=权威记录
  schema/{op-request,error,capability}.schema.json
  catalog/01-cell-range-io.md                # 唯一已写的域
languages/
  rust/   src/{com,excel}.rs(COM 层) ops/{mod,dispatch,catalog}.rs  bin/{excel-ops,capctl}.rs  main.rs(E01-E12) tests/
  openxml/ ExcelOps.csproj src/{Program,Contract,ExcelError,A1,Ops}.cs tests/  README.md
  cpp/    src/{main,tasks,excel_com}.* build.bat   # E01-E12 参考,含可复用 COM ExcelApp
  vba/    src/*.bas                                 # E01-E12 参考
```

---

## 4. 构建 / 测试 / 运行

```bash
# Rust(3 个 bin:excel-com=E01-E12, excel-ops=后端+参考读取器, capctl=编排器)
cd languages/rust && cargo build --bins
cargo test --lib            # 快,纯逻辑
cargo test                  # 含集成测试,会起 Excel,慢(~1min)

# OpenXML
dotnet build languages/openxml -c Debug      # 出 bin/Debug/net10.0/ExcelOps.exe
dotnet test languages/openxml                # 24 测试,用 OpenXML 回读,不起 Excel

# C++(E01-E12 参考;新后端可复用其 COM 层)
languages/cpp/src/build.bat                  # vcvarsall x64 + cl.exe

# 跑某后端的能力验证 + 回填矩阵(从 languages/rust/ 下):
cargo run --bin capctl -- verify \
  --backend openxml \
  --backend-cmd ../../languages/openxml/bin/Debug/net10.0/ExcelOps.exe \
  --reference-cmd target/debug/excel-ops.exe \
  --catalog ../../spec/capabilities/catalog \
  --out ../../spec/capabilities/support-matrix.md
```

---

## 5. 下一步任务(按优先级)

### 任务 A — C++ 后端(方案 #3)
- 新建一个 C++ 可执行(建议 `languages/cpp/src/backend/` 或并列 `excel_ops.cpp`),实现 `docs/backend-protocol.md` 的 stdin/stdout JSON 契约。
- **复用现成的 `languages/cpp/src/excel_com.{h,cpp}` 的 `ExcelApp`**(COM 封装,E01-E12 已验证可构建)。只需实现写类 op(cell.write/write_bulk/clear/copy_values)。
- 需要一个 C++ JSON 库;仓库目前没有。可用单头文件 `nlohmann/json`(放 `languages/cpp/src/third_party/json.hpp`)或自己解析。决定前先看 build.bat 怎么加文件。
- 验证:`capctl verify --backend cpp --backend-cmd <cpp exe> --reference-cmd <rust excel-ops> ...` → 回填 C++ 列。预期 Cell I/O 全 ✅。
- **先 brainstorm 一下 JSON 库选型 + exe 入口**,再 writing-plans。

### 任务 B — VBA 后端(方案 #4)—— 需要设计
- VBA 跑在 Excel 宿主里,**没有"独立 exe"**。要把它包成一个"读 stdin JSON / 写 stdout JSON"的命令,可行路径:一个 PowerShell/cscript 包装器,用 COM 起 Excel、注入并 `Run` 一段 VBA(参考 E11 的注入手法),VBA 内完成写操作。这是真正需要 brainstorm 的一项(宿主、信任设置、把 OpRequest 传给 VBA 的方式)。
- 注意:本机"信任访问 VBA 工程对象模型"默认**关闭**(E11 因此 SKIP)。VBA 后端要么要求开启该信任,要么用别的注入方式。

### 任务 C — 按域滚动扩充能力(用户的核心诉求)
- 现在已有 ≥2 个后端(Rust+OpenXML)。用户明确要求:**每新增一条能力,立即在所有可用后端实现 + 各自验证 + 回填一行**(批的单位 = 能力条目,不是语言)。
- 下一个域按设计稿顺序:**②格式化**(字体粗/斜/下划线/号/名/色/删除线、填充、边框、数字格式、对齐、换行、合并 ~30 条)→ ③行列结构 → ④公式 → ⑤表/簿 → ⑥数据 → ⑦图表 → ⑧IO/导出。
- 每条:写 catalog 条目(带 verify 配方)→ 在每个后端的 dispatch 里实现对应 op → `capctl verify` 各后端回填。**格式化域会暴露 OpenXML 的强项**(它擅长写样式),而**公式/图表/PDF 域会暴露 OpenXML 的 ❌**(无引擎)——这正是对比看点。
- 新增 op 时注意:capctl 的断言比较器目前**只支持标量**(number/string/bool/null);要断言 2D 数组需先扩 `values_match`(见已知 TODO)。

---

## 6. 已知 TODO / 推迟项(评审中识别,未做)

- **`range.read_bulk`(数组回读)**:Rust `com.rs` 的 `Variant::from_raw` 把 SAFEARRAY 映射成 `Unknown`,未实现数组读;且 capctl 比较器只比标量。要做整行/整块读回需补这两处。当前所有断言走单格读规避。
- **capctl 多后端聚合**:已实现"合并保留其它列"(读现有 matrix + 只更新本次 `--backend` 列)。够用。
- **OpenXML 小瑕疵**(非阻塞):`range.clear` 后会留空 `<row>` 元素(Excel 容忍);`range.copy_values` 源传带 `:` 的区间会报 RangeParseError(信息略隐晦);`clear` 的 `all` 与 `contents` 行为相同(删整个 `<c>`,未单独保样式)。
- **`CELL-READ-TEXT` 区域设置敏感**:本机英文 locale 下 ✅;非英文机器可能 ⚠️(已在 matrix/README 注明)。
- **Rust `com.rs` dispatch 实参引用泄漏**:作为 COM 实参传入的 IDispatch 会 AddRef 不释放(E08 Sort key / E09 SetSourceData range);E12 靠强杀进程兜底,12/12 仍 PASS。可日后清。
- **`save_as.format` 的 `pdf`** 已从 schema 移除(导出 PDF 将来作专门 op,不作为保存格式)。

---

## 7. 本次踩过的坑(给接手的人省时间)

- **vtMissing**:COM 可选参数必须传 `VT_ERROR`+`DISP_E_PARAMNOTFOUND`,**不能传 `VT_EMPTY`**,否则 Excel 报 0x800A03EC(1004)。Rust 侧 `com.rs` 的 `empty()` 已经是 `Variant::Missing`(=vtMissing)。
- **路径分隔符**:正斜杠路径会让 COM `SaveAs`/`Open` 失败(HRESULT -2146827284);Rust 侧 `win_path()` 统一转反斜杠。新后端(C++/VBA)注意同样问题。
- **OpenXML 单元格必须按行升序、行内按列升序插入**,否则文件"损坏"。`A1.cs` 的 `EnsureCell` 已处理。字符串用 **inline string**(`t="inlineStr"`)避开共享字符串表的复杂度。
- **OpenXML 公式**:写 `<f>2+3</f>` **不写 `<v>`**(不伪造缓存值);Excel 打开时自动重算 → 参考读取器读到结果。所以 `CELL-WRITE-FORMULA` OpenXML 也是 ✅(诚实测量的真实结果)。
- **依赖版本**:`DocumentFormat.OpenXml` 锁 `16.0.13823.15015`;NU1701 警告(包标 .NETFramework)无害;xUnit 锁 `2.9.3`(v3 runner 与当前 test platform 不兼容)。
- **.gitignore 根有 `bin/` 规则**(为 .NET 产物):会忽略任意层级的 `bin/` 目录。`languages/rust/src/bin/` 已加 `!` 反忽略;`.NET` 的 `bin/`/`obj/` 正确忽略(只跟踪 .cs/.csproj)。
- **serde**:`OpResponse` 用结构体 `{ok:bool, result?, error?}`,不是 `#[serde(tag="ok")]` 内部标签(布尔不能做 serde 内部 tag)。
- **Excel 新建工作簿默认 1 个 sheet**(本机);某些任务依赖此(E07 加表后期望 2)。
- **参考读取器是 Rust**:若 Rust 写+读共享同一个系统性 bug,会显示假 ✅。已在 backend-protocol §7 注明。

---

## 8. 约定 / 流程

- **分支**:从 `main` 切 `user/wuyin/<kebab-name>`;不在 `main` 上直接改。
- **commit 结尾**:`Co-Authored-By: Claude Opus 4.8 <noreply@anthropic.com>`。
- **可用技能**:`brainstorming`(动手前)、`writing-plans`、`subagent-driven-development`、`finishing-a-development-branch`、`commit-and-push`。
- **流程提醒**:上一手 agent 给每个 chunk 都套了"实现+spec审查+质量审查+修复"四道子代理,很稳但**慢**。用户明确希望**更轻的流程**(实现 + 跑测试验证为主,减少来回评审)。请据此把握节奏。
- **分支清理**:`user/wuyin/rust-excel-com-tasks`、`user/wuyin/openxml-backend` 都已 merge 进 `main`,本地/远程仍在,可删可留。
