# Handoff — Excel 自动化

> **Date:** 2026-06-09
> **Repo:** `XiaoYingGee/office-com-automation`

---

## 0. 当前状态

Excel 自动化 4-backend 实现完成：pywin32 / VBA / Python Add-in / C# Add-in。
前两个已跑完 benchmark（VBA 全面领先）。后两个 add-in 已实现，待注册后验证。

---

## 1. 前置条件

- Windows + Microsoft Excel (含 VBA)
- Python 3.10+ with `pywin32` (`pip install pywin32`) + `openpyxl`
- .NET SDK (for C# add-in build): `dotnet build -c Release`
- Excel Trust Center: 勾选 "Trust access to the VBA project object model"

---

## 2. Add-in 注册

### Python Add-in

```powershell
python python_addin\excel_pyaddin.py          # 注册
python python_addin\excel_pyaddin.py --unregister  # 卸载
# 重启 Excel 后生效
```

### C# Add-in

```powershell
cd csharp_addin\ExcelEditorAddin
dotnet build -c Release
cd ..
powershell .\register.ps1       # 注册
powershell .\unregister.ps1     # 卸载
# 重启 Excel 后生效
```

### 验证 Add-in 已加载

```python
import win32com.client
app = win32com.client.Dispatch("Excel.Application")
# Python add-in
bridge = app.COMAddIns.Item("ExcelEditor.PyAddIn").Object
print(bridge.Ping())  # → "pong"
# C# add-in
bridge = app.COMAddIns.Item("ExcelEditor.AddIn").Object
print(bridge.Ping())  # → "pong"
```

---

## 3. 使用

```powershell
cd skills\excel-editor\scripts

# 基本操作
python excel_editor.py test.xlsx --create --inspect
python excel_editor.py test.xlsx --inspect --backend vba

# 4-backend benchmark
python benchmark.py --all --size empty,medium,large --rounds 3

# 只跑 add-in
python benchmark.py --backends pyaddin,csharp-addin --size empty
```

---

## 4. 项目结构

```
office-com-automation/
├── HANDOFF.md
├── benchmark.md                  # 性能对比报告
├── docs/                         # COM 概述、Excel 对象模型
├── python_addin/
│   └── excel_pyaddin.py          # Python in-process COM add-in
├── csharp_addin/
│   ├── ExcelEditorAddin/
│   │   ├── Connect.cs            # C# add-in + action dispatch
│   │   └── ExcelEditorAddin.csproj
│   ├── register.ps1              # Per-user 注册（无需 admin）
│   └── unregister.ps1
└── skills/
    └── excel-editor/
        ├── SKILL.md
        ├── scripts/
        │   ├── excel_editor.py   # pywin32 + VBA engine
        │   ├── benchmark.py      # 4 backends × 11 tests × 3 sizes
        │   └── gen_fixture.py
        └── references/
            └── ExcelEditorBridge.bas
```

---

## 5. 架构（2×2 矩阵）

```
                  Out-of-Process              In-Process (零 IPC)
Python            pywin32 (baseline)          Python COM Add-in
VBA/C#            —                           VBA / C# COM Add-in
```

Add-in 驱动协议：外部 Python 通过 `COMAddIns.Item(progid).Object` 获取桥接对象，
每个操作只做 **1 次跨进程调用**（传 JSON 进去），所有 per-cell 遍历在 Excel 进程内完成。

---

## 6. 踩坑

- **路径分隔符**: 正斜杠让 COM 失败；`win_path()` 转反斜杠
- **VBA StringBuilder**: `SbAppend` + `Mid$` 做 O(n)
- **Add-in 只在交互式启动的 Excel 中加载**: 通过 Dispatch() 自动化启动的 Excel 不加载 COMAddIns
- **Per-user COM 注册**: HKCU\Software\Classes\CLSID，不需要 admin

---

## 7. 约定

- **Commit 结尾**: `Co-Authored-By: Claude Opus 4 (1M context) <noreply@anthropic.com>`
- **流程**: 轻量（实现 + 跑测试验证）
- **直接在 main 上工作**
