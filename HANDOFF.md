# Handoff — Excel 自动化 Benchmark

> **Date:** 2026-06-09
> **Repo:** `XiaoYingGee/office-com-automation`
> **接手方式:** 读本文件，在 Windows 上跑 benchmark。

---

## 0. 当前目标

**在 Windows 机器上跑 pywin32 / VBA benchmark，出对比数据。**

两个后端(pywin32 / VBA)已实现统一的 11 个 op，benchmark 脚本覆盖 9 项测试。代码已就绪。

---

## 1. 前置条件

- Windows + Microsoft Excel (含 VBA)
- Python 3.10+ with `pywin32` (`pip install pywin32`) + `openpyxl` (`pip install openpyxl`，用于生成 benchmark fixture)
- Excel Trust Center: 勾选 "Trust access to the VBA project object model"（注册表 `HKCU\Software\Microsoft\Office\16.0\Excel\Security\AccessVBOM = 1`）

---

## 2. 运行步骤

```powershell
# 1. 验证 skill 工作
cd skills\excel-editor\scripts
python excel_editor.py test.xlsx --create --inspect
python excel_editor.py test.xlsx --inspect --backend vba

# 2. 跑完整 benchmark
python benchmark.py --backends pywin32,vba --size empty,medium,large --rounds 3

# 3. 查看结果
type ..\..\..\benchmarks\results\benchmark-results.md
```

---

## 3. 项目结构

```
office-com-automation/
├── HANDOFF.md                    # ★ 本文件
├── docs/                         # COM 概述、对象模型
├── spec/                         # 标准任务集 + capability matrix
├── languages/
│   └── vba/                      # VBA E01-E12 参考 (.bas)
├── skills/
│   └── excel-editor/             # ★ Claude Code skill
│       ├── SKILL.md
│       ├── scripts/
│       │   ├── excel_editor.py   # Python 入口 (pywin32 + VBA backend)
│       │   ├── benchmark.py      # benchmark (2 backends × 11 tests)
│       │   └── gen_fixture.py    # 生成 empty/medium/large fixture
│       └── references/
│           └── ExcelEditorBridge.bas  # VBA 桥接宏
└── benchmarks/results/           # benchmark 输出目录
```

---

## 4. 统一 Op 集 (13 个)

| Op | pywin32 | VBA |
|----|:---:|:---:|
| cell.write | ✅ | ✅ |
| range.write_bulk | ✅ | ✅ |
| range.read | ✅ | ✅ |
| range.clear | ✅ | ✅ |
| range.merge | ✅ | ✅ |
| range.unmerge | ✅ | ✅ |
| range.copy_values | ✅ | ✅ |
| inspect | ✅ | ✅ |
| set_format | ✅ | ✅ |
| row.insert | ✅ | ✅ |
| row.delete | ✅ | ✅ |
| sheet.add | ✅ | ✅ |
| sheet.rename | ✅ | ✅ |
| sheet.delete | ✅ | ✅ |

---

## 5. Benchmark（11 项 × 3 尺寸）

测试 B0–B10（含 B0 Open、B10 Merge），跨 empty / medium / large 三种文档规模。
完整结果见根目录 `benchmark.md`。

### 最终结论 (2026-06-09)

**VBA 几乎全胜**，IPC 密集型差距最大（medium 的 Batch 36×、Merge 51×、Inspect 21×）；
pywin32 仅在打开文件和首次写略快。大文档代价集中在**一次性打开**（large ~18.5s 开 145MB），
打开后单次操作吞吐与小文档相当。

---

## 6. 踩坑

- **vtMissing**: COM 可选参数必须传 `VT_ERROR`+`DISP_E_PARAMNOTFOUND`
- **路径分隔符**: 正斜杠让 COM 失败；`win_path()` 转反斜杠
- **VBA StringBuilder**: `SbAppend` + `Mid$` 做 O(n)
- **VBA 引号**: 用 `Chr(34)` 避免编译器死循环
- **benchmark.py 默认路径**: 脚本在 `skills/excel-editor/scripts`，到 repo 根需 **3 层** `..`（早期 2 层 bug 把结果写进了 `skills/benchmarks/`）

---

## 7. 约定

- **Commit 结尾**: `Co-Authored-By: Claude Opus 4 (1M context) <noreply@anthropic.com>`
- **流程**: 轻量（实现 + 跑测试验证）
- **直接在 main 上工作**
