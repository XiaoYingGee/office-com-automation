# Handoff — Excel Skill + 4-Backend Benchmark

> **Date:** 2026-06-09
> **Branch:** `user/wuyin/excel-skill-benchmark`
> **Status:** Code complete, needs Windows testing

---

## 0. 一句话现状

新增了一个 **Excel COM 自动化 skill**（`skills/excel-editor/`），并将 Rust 和 OpenXML 后端扩展到与 pywin32/VBA 相同的 op 集，使得 4 个后端可以跑统一的 9 项 benchmark。

---

## 1. 本次变更

### 新增: Excel Editor Skill (`skills/excel-editor/`)

```
skills/excel-editor/
├── SKILL.md                          # Skill 定义 + 用法文档
├── scripts/
│   ├── excel_editor.py               # Python 入口 (ExcelCOM + ExcelVBA 两个 backend class)
│   └── benchmark.py                  # 统一 benchmark 脚本 (4 backends × 9 tests)
└── references/
    └── ExcelEditorBridge.bas         # VBA 进程内桥接宏 (StringBuilder JSON)
```

- **pywin32 backend**: Python 直接通过 win32com 操作 Excel COM（基线，每属性一次 IPC）
- **VBA backend**: Python 通过 `Application.Run` 调用注入的 VBA 宏（进程内，零 IPC）
- VBA 模块使用 `StringBuilder` 模式（O(n) `Mid$` 写入），不是 O(n²) 字符串拼接

### 扩展: Rust Backend (`languages/rust/src/ops/dispatch.rs`)

新增 7 个 op:
- `inspect` — 返回工作簿结构 JSON
- `set_format` — 格式化区域（bold, italic, font_size, font_name, font_color, bg_color, number_format）
- `row.insert` / `row.delete` — 插入/删除行
- `sheet.add` / `sheet.rename` / `sheet.delete` — 工作表管理

### 扩展: OpenXML Backend (`languages/openxml/src/Ops.cs`)

新增 8 个 op:
- `range.read` — 读取单元格值
- `inspect` — 工作簿结构
- `set_format` — 格式化（通过 Stylesheet 操作）
- `row.insert` / `row.delete` — 行操作（含 cell reference 更新）
- `sheet.add` / `sheet.rename` / `sheet.delete` — 工作表管理

---

## 2. 4 个后端统一 Op 集

| Op | Rust (COM) | OpenXML (.NET) | pywin32 | VBA |
|----|:---:|:---:|:---:|:---:|
| cell.write | ✅ | ✅ | ✅ | ✅ |
| range.write_bulk | ✅ | ✅ | ✅ | ✅ |
| range.read | ✅ | ✅ | ✅ | ✅ |
| range.clear | ✅ | ✅ | ✅ | ✅ |
| inspect | ✅ | ✅ | ✅ | ✅ |
| set_format | ✅ | ✅ | ✅ | ✅ |
| row.insert | ✅ | ✅ | ✅ | ✅ |
| row.delete | ✅ | ✅ | ✅ | ✅ |
| sheet.add | ✅ | ✅ | ✅ | ✅ |
| sheet.rename | ✅ | ✅ | ✅ | ✅ |
| sheet.delete | ✅ | ✅ | ✅ | ✅ |

---

## 3. 如何运行 Benchmark

### 前置条件

- Windows + Microsoft Excel (含 VBA)
- Python 3.10+ with `pywin32` (`pip install pywin32`)
- Rust toolchain (`cargo build --release`)
- .NET SDK 9.0+ (`dotnet build -c Release`)
- **VBA backend 需要**: File → Options → Trust Center → Trust Center Settings → Macro Settings → 勾选 "Trust access to the VBA project object model"

### 构建

```powershell
# Rust
cd languages\rust
cargo build --release

# OpenXML
cd languages\openxml
dotnet build -c Release
```

### 运行 benchmark

```powershell
cd skills\excel-editor\scripts

# 只测 pywin32 + VBA (无需编译 Rust/OpenXML)
python benchmark.py --backends pywin32,vba --rounds 3

# 全部 4 个 backend
python benchmark.py --all ^
  --rust-exe ..\..\..\languages\rust\target\release\excel-ops.exe ^
  --openxml-exe ..\..\..\languages\openxml\bin\Release\net10.0\ExcelOps.exe ^
  --rounds 5

# 带可见 Excel 窗口 (调试用)
python benchmark.py --backends pywin32 --headed --rounds 1
```

### 输出

- `benchmarks/results/benchmark-results.md` — Markdown 表格
- `benchmarks/results/benchmark-results.json` — 原始数据

### 单独测试 skill

```powershell
cd skills\excel-editor\scripts

# 创建新文件
python excel_editor.py test.xlsx --create --inspect

# 执行 actions
python excel_editor.py test.xlsx --exec-actions "[{\"action\":\"write_cell\",\"sheet\":\"Sheet1\",\"cell\":\"A1\",\"value\":42}]"

# VBA backend
python excel_editor.py test.xlsx --inspect --backend vba

# 交互模式
python excel_editor.py test.xlsx --interactive-actions --backend vba
```

---

## 4. 预期结果 & 假设

核心假设（与 PPT 项目一致）：**IPC 调用次数决定性能**

| Backend | 架构 | 预期性能排名 |
|---------|------|:---:|
| **VBA** | 1 次 App.Run → 进程内执行 | 🥇 |
| **pywin32** | 每属性 1 次跨进程 COM IPC | 🥈 (inspect 会很慢) |
| **OpenXML** | 文件 I/O，无 Excel 进程 | 🥉 (小文件快，但无法 recalc) |
| **Rust** | 无状态，每 op spawn Excel → 操作 → 关闭 | 🏅 (进程启动开销巨大) |

关键看点:
- **B5 (inspect)**: VBA 应远快于 pywin32（零 IPC vs N×IPC）
- **B7 (set_format)**: 格式化需多次属性设置，IPC 差异放大
- **B2 (bulk write)**: 批量写入用 SAFEARRAY 一次传输，所有 COM backend 差距缩小
- **Rust**: 因为每个 op 都 spawn 新 Excel 进程，总耗时 = (N ops × ~2s spawn) >> 其它 backend

---

## 5. 已知待验证项

- [ ] Rust 新 op 是否能编译（需 Windows + `windows` crate target）
- [ ] OpenXML `set_format` 的 Stylesheet 操作是否正确（新建 Font/Fill/CellFormat 索引）
- [ ] VBA 注入是否正常工作（需"信任 VBA 工程对象模型"开启）
- [ ] `row.insert` 在 Rust 里用 `ws.Rows.Item(i).Insert`，需确认 COM dispatch 路径
- [ ] OpenXML `row.insert` 的 cell reference 更新逻辑是否完整
- [ ] Benchmark 脚本在 Windows 路径下的运行（反斜杠 / 临时目录）

---

## 6. 约定

- 分支: `user/wuyin/excel-skill-benchmark`（从 main HEAD 切出）
- 测试通过后 merge 到 main
- Commit 风格: `feat:` / `fix:`
