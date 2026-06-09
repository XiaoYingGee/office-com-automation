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
- Python 3.10+ with `pywin32` (`pip install pywin32`) + `openpyxl` (`pip install openpyxl`，用于生成 benchmark fixture)
- Rust toolchain (`cargo build --release`)
- .NET SDK 10.0 (`dotnet build -c Release`，输出 `bin/Release/net10.0/`)
- C++: MSVC (VS 2022 BuildTools，`languages/cpp/src/build.bat` 用 vcvarsall + cl 编译 → `excel-ops-cpp.exe`)
- Excel Trust Center: 勾选 "Trust access to the VBA project object model"（注册表 `HKCU\Software\Microsoft\Office\16.0\Excel\Security\AccessVBOM = 1`）

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

## 4. 统一 Op 集 (13 个)

| Op | Rust (COM) | OpenXML (.NET) | C++ (COM) | pywin32 | VBA |
|----|:---:|:---:|:---:|:---:|:---:|
| cell.write | ✅ | ✅ | ✅ | ✅ | ✅ |
| range.write_bulk | ✅ | ✅ | ✅ | ✅ | ✅ |
| range.read | ✅ | ✅ | ✅ | ✅ | ✅ |
| range.clear | ✅ | ✅ | ✅ | ✅ | ✅ |
| range.merge | ✅ | ✅ | ✅ | ✅ | ✅ |
| range.unmerge | ✅ | ✅ | ✅ | ✅ | ✅ |
| range.copy_values | ✅ | ✅ | ✅ | ✅ | ✅ |
| inspect | ✅ | ✅ | ✅ | ✅ | ✅ |
| set_format | ✅ | ✅ | ✅ | ✅ | ✅ |
| row.insert | ✅ | ✅ | ✅ | ✅ | ✅ |
| row.delete | ✅ | ✅ | ✅ | ✅ | ✅ |
| sheet.add | ✅ | ✅ | ✅ | ✅ | ✅ |
| sheet.rename | ✅ | ✅ | ✅ | ✅ | ✅ |
| sheet.delete | ✅ | ✅ | ✅ | ✅ | ✅ |

---

## 5. Benchmark（11 项 × 3 尺寸）

测试 B0–B10（含 B0 Open、B10 Merge），跨 empty / medium / large 三种文档规模。
完整结果见根目录 `benchmark.md`。

### 最终结论 (2026-06-09)

最终对比聚焦两个**常驻会话**后端：**VBA 几乎全胜**，IPC 密集型差距最大
（medium 的 Batch 36×、Merge 51×、Inspect 21×）；pywin32 仅在打开文件和首次写略快。
大文档代价集中在**一次性打开**（large ~18.5s 开 145MB），打开后单次操作吞吐与小文档相当。

**无状态后端（Rust / OpenXML / C++）已实现并验证全部 13 op，但因每 op 重开整个文件、
比常驻后端慢 1–2 个数量级、大文档完全不可行，已从最终对比移除**（代码保留在 `languages/`）。
详见 `benchmark.md` §6。

---

## 6. 踩坑

- **vtMissing**: COM 可选参数必须传 `VT_ERROR`+`DISP_E_PARAMNOTFOUND`
- **路径分隔符**: 正斜杠让 COM 失败；`win_path()` 转反斜杠
- **OpenXML 行列排序**: 单元格必须按行升序插入
- **VBA StringBuilder**: `SbAppend` + `Mid$` 做 O(n)
- **VBA 引号**: 用 `Chr(34)` 避免编译器死循环
- **C++ `GetAddress`**: `#import` 生成的签名前 3 参必填 → `GetAddress(vtMissing, vtMissing, Excel::xlA1)`
- **C++ 工具链**: cl.exe 在 VS2022 **BuildTools**（`C:\Program Files (x86)\...\BuildTools`），不是 Enterprise；`build.bat` 经 vcvarsall 调用
- **benchmark.py 默认路径**: 脚本在 `skills/excel-editor/scripts`，到 repo 根需 **3 层** `..`（早期 2 层 bug 把结果写进了 `skills/benchmarks/`）
- **Bash 工具传 JSON**: 反斜杠路径会被吞 → exe 路径用正斜杠绝对路径

---

## 7. 约定

- **Commit 结尾**: `Co-Authored-By: Claude Opus 4 (1M context) <noreply@anthropic.com>`
- **流程**: 轻量（实现 + 跑测试验证）
- **直接在 main 上工作**
