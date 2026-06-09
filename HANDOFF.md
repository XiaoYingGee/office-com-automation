# Handoff — Excel 自动化

> **Date:** 2026-06-09
> **Repo:** `XiaoYingGee/office-com-automation`

---

## 0. 当前状态

Excel 自动化 skill 已完成，支持 pywin32 和 VBA 两个后端。Benchmark 已跑完，VBA 全面领先。

---

## 1. 前置条件

- Windows + Microsoft Excel (含 VBA)
- Python 3.10+ with `pywin32` (`pip install pywin32`) + `openpyxl` (`pip install openpyxl`)
- Excel Trust Center: 勾选 "Trust access to the VBA project object model"（注册表 `HKCU\Software\Microsoft\Office\16.0\Excel\Security\AccessVBOM = 1`）

---

## 2. 使用

```powershell
cd skills\excel-editor\scripts
python excel_editor.py test.xlsx --create --inspect
python excel_editor.py test.xlsx --inspect --backend vba

# benchmark
python benchmark.py --backends pywin32,vba --size empty,medium,large --rounds 3
```

---

## 3. 项目结构

```
office-com-automation/
├── HANDOFF.md                    # ★ 本文件
├── benchmark.md                  # pywin32 vs VBA 性能对比报告
├── docs/                         # COM 概述、Excel 对象模型
└── skills/
    └── excel-editor/             # ★ Claude Code skill
        ├── SKILL.md
        ├── scripts/
        │   ├── excel_editor.py   # Python 入口 (pywin32 + VBA backend)
        │   ├── benchmark.py      # benchmark (2 backends × 11 tests × 3 sizes)
        │   └── gen_fixture.py    # 生成 empty/medium/large fixture
        └── references/
            └── ExcelEditorBridge.bas  # VBA 桥接宏
```

---

## 4. Op 集 (13 个)

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

## 5. 踩坑

- **路径分隔符**: 正斜杠让 COM 失败；`win_path()` 转反斜杠
- **VBA StringBuilder**: `SbAppend` + `Mid$` 做 O(n)
- **VBA 引号**: 用 `Chr(34)` 避免编译器死循环
- **benchmark.py 默认路径**: 脚本在 `skills/excel-editor/scripts`，到 repo 根需 **3 层** `..`

---

## 6. 约定

- **Commit 结尾**: `Co-Authored-By: Claude Opus 4 (1M context) <noreply@anthropic.com>`
- **流程**: 轻量（实现 + 跑测试验证）
- **直接在 main 上工作**
