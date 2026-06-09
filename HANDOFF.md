# Handoff — Excel 自动化 4-Backend Benchmark

> **Date:** 2026-06-09
> **Repo:** `XiaoYingGee/office-com-automation`
> **接手方式:** 读本文件，在 Windows 上跑 benchmark。

---

## 0. 当前目标

**在 Windows 机器上跑 4-backend benchmark，出对比数据。**

4 个后端(Rust COM / OpenXML .NET / pywin32 / VBA)已实现统一的 11 个 op，benchmark 脚本覆盖 9 项测试。代码已就绪，需要 Windows + Excel 环境验证编译和运行。

---

## 1. 前置条件

- Windows + Microsoft Excel (含 VBA)
- Python 3.10+ with `pywin32` (`pip install pywin32`)
- Rust toolchain (`cargo build --release`)
- .NET SDK 9.0+ (`dotnet build -c Release`)
- Excel Trust Center: 勾选 "Trust access to the VBA project object model"

---

## 2. 运行步骤

```powershell
# 1. 构建 Rust backend
cd languages\rust
cargo build --release

# 2. 构建 OpenXML backend
cd languages\openxml
dotnet build -c Release

# 3. 验证 skill 工作
cd skills\excel-editor\scripts
python excel_editor.py test.xlsx --create --inspect
python excel_editor.py test.xlsx --inspect --backend vba

# 4. 跑完整 benchmark
python benchmark.py --all ^
  --rust-exe ..\..\..\languages\rust\target\release\excel-ops.exe ^
  --openxml-exe ..\..\..\languages\openxml\bin\Release\net10.0\ExcelOps.exe ^
  --rounds 5

# 5. 查看结果
type ..\..\..\benchmarks\results\benchmark-results.md
```

---

## 3. 项目结构

```
office-com-automation/
├── HANDOFF.md                    # ★ 本文件
├── docs/                         # COM 概述、对象模型、backend-protocol 契约
├── spec/                         # 能力目录 + capability matrix
├── languages/
│   ├── rust/                     # Rust COM backend (excel-ops.exe)
│   ├── openxml/                  # .NET OpenXML backend (ExcelOps.exe)
│   ├── cpp/                      # C++ COM E01-E12
│   └── vba/                      # VBA E01-E12 参考 (.bas)
├── skills/
│   └── excel-editor/             # ★ Claude Code skill
│       ├── SKILL.md
│       ├── scripts/
│       │   ├── excel_editor.py   # Python 入口 (pywin32 + VBA backend)
│       │   └── benchmark.py      # 统一 benchmark (4 backends × 9 tests)
│       └── references/
│           └── ExcelEditorBridge.bas  # VBA 桥接宏
└── benchmarks/results/           # benchmark 输出目录
```

---

## 4. 统一 Op 集 (11 个)

| Op | Rust (COM) | OpenXML (.NET) | pywin32 | VBA |
|----|:---:|:---:|:---:|:---:|
| cell.write | ✅ | ✅ | ✅ | ✅ |
| range.write_bulk | ✅ | ✅ | ✅ | ✅ |
| range.read | ✅ | ✅ | ✅ | ✅ |
| range.clear | ✅ | ✅ | ✅ | ✅ |
| range.copy_values | ✅ | ✅ | ✅ | ✅ |
| inspect | ✅ | ✅ | ✅ | ✅ |
| set_format | ✅ | ✅ | ✅ | ✅ |
| row.insert | ✅ | ✅ | ✅ | ✅ |
| row.delete | ✅ | ✅ | ✅ | ✅ |
| sheet.add | ✅ | ✅ | ✅ | ✅ |
| sheet.rename | ✅ | ✅ | ✅ | ✅ |
| sheet.delete | ✅ | ✅ | ✅ | ✅ |

---

## 5. Benchmark 9 项测试

| # | 测试 | 考察点 |
|---|------|--------|
| B1 | Cell Write (5 cells) | 基础写入延迟 |
| B2 | Bulk Write (100×10) | 批量写入 |
| B3 | Read Cell | 读取延迟 |
| B4 | Clear Range (1000 cells) | 批量清除 |
| B5 | Inspect Workbook | 读密集(IPC差异最大) |
| B6 | Batch 10 Writes | 多次操作累积 |
| B7 | Format 5 Cells | 多属性设置 |
| B8 | Insert 5 Rows | 结构操作 |
| B9 | Sheet Add+Rename+Delete | 工作表管理 |

### 实测排名 (2026-06-09, 3/4 backends)

> 详见 `benchmarks/results/benchmark-results.md`。Rust 未测 (每 op spawn Excel，micro-op 套件下过慢)。

| Backend | 架构 | 预期 | **实测** |
|---------|------|:---:|:---:|
| **VBA** | Python → App.Run(1次) → 进程内执行 | 🥇 | 🥇 (8/9 最快) |
| **pywin32** | Python → 每属性1次跨进程COM IPC | 🥈 | 🥈 |
| **OpenXML** | standalone exe，每 op 启动进程+整文件读写 | 🥉 | **最慢** (每 op ~300ms 固定开销) |
| **Rust** | 无状态: 每op spawn Excel→操作→关闭 | 4th | 未测 |

**⚠️ 预期修正**: OpenXML 原预测 🥉/"快速文件 I/O"，实测是所有测试里最慢的。
原因: standalone exe **每 op** 都 *启动进程 + 读整个 .xlsx + 改 + 整文件写回*，
~300ms 固定开销压倒"不启动 Excel"的收益。其优势仅在大批量一次性变换场景才体现，
本 benchmark 的小粒度高频 op 反而是其最差场景。

---

## 6. 待验证

- [ ] Rust 新 op 编译通过
- [ ] OpenXML set_format Stylesheet 操作正确
- [ ] VBA 注入正常（需 Trust Center 开启）
- [ ] Benchmark Windows 上完整跑通
- [ ] row.insert COM dispatch 路径确认

---

## 7. 踩坑

- **vtMissing**: COM 可选参数必须传 `VT_ERROR`+`DISP_E_PARAMNOTFOUND`
- **路径分隔符**: 正斜杠让 COM 失败；`win_path()` 转反斜杠
- **OpenXML 行列排序**: 单元格必须按行升序插入
- **VBA StringBuilder**: `SbAppend` + `Mid$` 做 O(n)
- **VBA 引号**: 用 `Chr(34)` 避免编译器死循环

---

## 8. 约定

- **Commit 结尾**: `Co-Authored-By: Claude Opus 4 (1M context) <noreply@anthropic.com>`
- **流程**: 轻量（实现 + 跑测试验证）
- **直接在 main 上工作**
