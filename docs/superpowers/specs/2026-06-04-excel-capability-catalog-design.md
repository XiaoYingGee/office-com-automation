# Excel 能力表 + 统一契约 — 设计稿

> Date: 2026-06-04
> Status: Approved (brainstorming)
> Topic: 语言无关的 Excel 能力目录 (capability catalog) + 统一操作契约，作为 4 个方案 (VBA/C++/Rust/OpenXML) 共同实现与横向 PK 的底座。

## 1. 背景与目标

本仓库是多语言 COM 自动化"比武场"：VBA / C++ / Rust 三语言实现同一组 Excel 动作，按四维 PK 对比。`E01–E12` 已三语言 12/12 PASS（含 Rust，本设计前一阶段完成）。

下一阶段的核心需求（来自 HANDOFF Expanded Scope，经 brainstorming 校准）：

> 广泛梳理一份"能力表"——Excel 中被普遍使用的功能，定义成语言无关的原子操作。这份表统一发包给三种语言实现，从而**看到每种语言对这些广泛功能的支持程度**。这是 PK"功能覆盖度"维度的细粒度底座。

**本设计的产物是这份能力表的结构、统一契约、自动验证模型与逐语言支持矩阵的规格定义**，以及随后的分批实现安排。

## 2. 关键决策（brainstorming 结论）

| 维度 | 决策 |
|---|---|
| 子项目定位 | 深度 API → 实为"语言无关能力表 + 契约"，三语言共用 |
| API 形态 | 意图导向的原子操作，入参/出参结构化、可序列化（为 MCP 留接缝） |
| 状态模型 | **无状态**：每个 op = 开→改→存→关；多步用 `batch.apply` 兜性能 |
| 错误模型 | 结构化 `{ category, code, message, hint }` |
| 覆盖广度 v1 | 核心 8 域做深；冷门特性缓后 |
| 能力表粒度 | **细**：拆到单属性，目标 ~120 项（80–150） |
| 支持度判定 | **自动验证**：每项带验证配方（写→存→**用统一参考读取器重开**→断言） |
| 实现方案 | **4 个方案**：VBA / C++ / Rust（均走 COM）+ **OpenXML**（.NET Open XML SDK，文件级、不起 Excel）。Office.js 暂不进（runtime 不契合 headless，记 roadmap）。 |
| 批的单位 | **能力条目**（非语言）：每新增一条能力，**立即在全部 4 方案实现 + 各自验证 + 回填一行**。前置一次性"基础设施批"先把契约与各方案骨架/验证 harness 立起来。 |

## 3. 交付物与文件布局

```
spec/capabilities/
├── README.md            # 总览 + 怎么读 + 支持度图例
├── contract.md          # 统一契约：OpRequest/OpResponse/错误模型/批量/无状态语义
├── catalog/             # 能力表本体，按域分文件
│   ├── 01-cell-range-io.md
│   ├── 02-formatting.md
│   ├── 03-structure.md
│   ├── 04-formula-calc.md
│   ├── 05-sheet-workbook.md
│   ├── 06-data-ops.md
│   ├── 07-charts.md
│   └── 08-io-export.md
├── schema/              # 机器可读 (供 MCP/校验器/PK 复用)
│   ├── op-request.schema.json
│   ├── error.schema.json
│   └── capability.schema.json
└── support-matrix.md    # 细粒度逐语言支持矩阵 (VBA/C++/Rust 列，实现后回填)
```

现有 `spec/capability-matrix.md`（12 任务粗粒度）保留，加指针指向这套细粒度矩阵。

## 4. 能力表结构（细粒度，~120 项）

每条能力固定字段：

| 字段 | 说明 |
|---|---|
| `id` | 稳定标识，如 `FMT-FONT-BOLD` |
| `name` / `desc` | 人读名称与描述 |
| `op` + `param_path` | 映射到契约里的哪个操作、哪个参数字段 |
| `sample` | 一组合法入参示例 |
| `verify` | 自动验证配方（见 §6） |
| `errors` | 该能力的非法入参 → 期望错误 category |
| `com_ref` | 底层 Excel 对象/属性，如 `Range.Font.Bold`（给三语言对齐） |
| `support` | VBA/C++/Rust/OpenXML 四列，实现后回填 ✅/⚠️/❌/⬜ |

**域与大致条目数（合计 ~120）：**

1. 单格/区域 I/O — ~15（按类型读写值、Value2/Text/Formula、批量区域读写、清除、复制值、地址解析）
2. 格式化 — ~30（字体 9：粗/斜/下划线/号/名/色/删除线/上下标；填充 3；边框 ~10：四边+对角 × 样式/粗细/色；数字格式；水平/垂直对齐；自动换行；缩进；方向；合并/取消）
3. 行列结构 — ~15（插/删行列、行高、列宽、自适应行/列、隐藏/取消、冻结窗格、分组/分级、打印区域）
4. 公式与计算 — ~10（设公式、数组公式、命名区域定义/删、计算模式手动/自动、强制重算、R1C1）
5. 工作表/簿 — ~15（增/删/改名/移动/复制表、标签色、激活、表数、保护/取消、可见性）
6. 数据操作 — ~12（单/多键排序、自动筛选设/清、高级筛选、查找、替换、去重、分列、数据验证列表）
7. 图表 — ~12（加图、源数据、类型、标题、轴标题、图例、数据标签、系列色、移动/尺寸、导出图片）
8. IO/导出 — ~12（新建、打开、保存、另存格式 xlsx/xlsm/csv/xls、导出 PDF、导出单表 PDF、关闭、文档属性）
9. （可选）宏 — ~3（注入 VBA 模块、运行宏）

**示例一条：**
```
id: FMT-FONT-COLOR
op: range.set_format   param_path: font.color
sample: { target: "Sheet1!A1", font: { color: "#FF0000" } }
verify: write sample → save → reopen → read range.format.font.color == "#FF0000"
errors: color="zzz" → InvalidArg ; target="A1:" → RangeParseError
com_ref: Range.Font.Color (BGR long)
support: { vba: ⬜, cpp: ⬜, rust: ⬜, openxml: ⬜ }
```

## 5. 统一契约 (contract.md)

- **无状态**：每个 op 自带 `path`，语义 = 打开→应用→保存→关闭。
- **OpRequest**：`{ op, path, target?: {sheet, range}, params: {...}, save_as?: {path, format} }`
- **OpResponse**：`{ ok: true, result: {...} }` 或 `{ ok: false, error: {...} }`
- **错误模型**：`{ category, code, message, hint }`，category 枚举：
  `FileNotFound | FileLocked | InvalidArg | RangeParseError | SheetNotFound | UnsupportedFormat | ComError | MacroTrustDisabled | Timeout | Unknown`
- **操作集较粗、参数化**（约 30–40 个 op）：细粒度能力通过 `param_path` 落到某 op 的某字段；如 `range.set_format` 一个 op 承载 font/fill/borders/number_format/alignment/merge 全部子属性。op 表面可控，比较表仍细。
- **批量**：`batch.apply { path, ops: [...] }` 在一次开/存/关里顺序套用，兜无状态多步编辑的性能短板。

## 6. 自动验证模型

每条能力的 `verify` 是声明式配方：
```
{ action: <op+params>, reopen: true|false, assert: { read: <op>, expect: <value>, tol?: <number> } }
```
实现方跑 action →（必要时重开）→读回→比对，**自动**判出 ✅/⚠️/❌。
- ✅ 完整：验证通过。
- ⚠️ 部分：能写入但重开丢失/有损/需 workaround（脚注说明）。
- ❌ 不支持：op 报错或结果错。

**统一参考读取器**：断言阶段的"重开读回"由**单一参考读取器（Excel COM）**完成——无论文件是哪个方案写的，都用同一把尺量。这样跨方案的判定口径一致；OpenXML 写出的 `.xlsx` 也用 COM 重开验证。各方案只负责"写"，"读回断言"统一。

`errors` 字段 → 失败路径验证，覆盖"边界/失败用例"目标。

## 7. 逐方案支持矩阵 (support-matrix.md)

行=能力 id，列=**VBA / C++ / Rust / OpenXML**，值=✅/⚠️/❌/⬜ + 脚注原因。
覆盖率 = (✅记1 + ⚠️记0.5) / 总数 → PK D1（功能覆盖度）得分。
OpenXML 因不起 Excel 引擎，预期在公式重算/图表渲染/PDF/宏等域大面积 ❌——这是高信息量的对比数据，体现"零 Office 依赖"的代价。

## 8. 分批实现安排（批的单位 = 能力条目）

- **基础设施批（前置一次性）**：契约 + schema + 校验器 + **统一参考读取器** + 4 方案各自的 op 骨架/验证 harness（Rust 已就绪先立，C++ 用现有 `vcvarsall+cl` 模式，OpenXML 新建 .NET 工程，VBA 走脚本宿主）。立完即可跑通"一条能力 → 4 方案实现 → 各自验证 → 回填一行"的端到端闭环。
- **能力滚动批（逐条目纵切）**：之后按域顺序逐条新增能力——每条**立即在 4 方案实现 + 各自自动验证 + 回填支持矩阵一行 + 提交**。不再按语言分批。
  - 顺序建议：① 单格/区域 I/O → ② 格式化 → ③ 结构 → ④ 公式 → ⑤ 表/簿 → ⑥ 数据 → ⑦ 图表 → ⑧ IO/导出。
  - 每条目以"4 方案各自验证可复现 + 矩阵该行填满（含 ❌ 脚注）"为完成口径。

> 写作计划拆分：基础设施批 = 第一份计划；能力滚动批按域切成后续若干份计划（每份一个域，逐条目跨 4 方案）。

## 9. 明确不做（YAGNI · v1）

- 冷门特性缓后：透视表、条件格式、数据验证(列表以外)、图片/形状、迷你图、切片器、批注、超链接、主题/单元格样式。
- **Office.js 不进本期**：其 runtime（云端 Office Scripts / 加载项宿主）不契合 headless COM/Wine 流水线，自动验证需另起一套 runner；记 roadmap 远期。
- 不做有状态 session、不做 MCP server（后续子项目；契约已留接缝）。
- Wine 运行时、四维跑分（走原 roadmap P2/P3，与本设计正交）。

## 10. 与现有结构的衔接

- 仓库框架小改：从"多语言 **COM** 比武"拓宽为"多**方案** Excel 自动化比武"（COM 引擎 vs OpenXML 文件级）。README/docs 相应措辞微调。
- `spec/capability-matrix.md`（12 任务粗矩阵）→ 加指针指向 `spec/capabilities/support-matrix.md`。
- `docs/pk-framework.md` D1 计分来源切换/补充为细粒度矩阵覆盖率。
- 不动 `languages/*/src` 现有 `E01–E12` 实现；新增能力实现各方案新开模块/入口（新增 `languages/openxml/`）。
