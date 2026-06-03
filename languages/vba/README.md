# VBA · 实现 Plan

> 参赛语言之一。Office 的"母语"——VBA 跑在 Office 进程**内部**，对象模型零封送、零外部运行时。
> 实际代码在 P1 写入本目录 `src/`。

## 定位

- **绑定**：早绑定，宿主内直接持有对象模型。
- **优势**：**无跨进程封送**（代码与 Excel 同进程）、无外部运行时依赖、写法最简洁、Wine 下只依赖 Office 本身。
- **劣势**：宿主受限（要在 Excel/Office 内运行）、工程化/版本管理弱、CI/headless 触发较别扭。

## 环境前提

- Windows + Excel（含 VBA，默认随 Office）。
- 宏安全：需允许运行宏；E11 注入场景需"信任对 VBA 工程对象模型的访问"。
- Wine：只依赖 Office 在 Wine 的可用性，预判相对乐观（🟢），P2 实测。

## 接入方式

VBA 在 Excel 内 `Application` 即当前实例；也可在独立宏工作簿里驱动新建实例：

```vba
' 宿主内：直接用 ActiveWorkbook / ThisWorkbook
Dim ws As Worksheet
Set ws = ThisWorkbook.Worksheets(1)

' 或新建后台实例（更贴近其它语言的可比口径）
Dim app As Excel.Application
Set app = New Excel.Application
app.Visible = False
```

> 为与其它语言**公平对比**，建议 VBA 也走"新建独立实例 → 操作 → 释放"的口径，而非只改 ActiveWorkbook。

## E01–E12 实现要点

| 任务 | 关键写法 / 坑 |
|---|---|
| E01 | `Workbooks.Add` → `SaveAs path, xlOpenXMLWorkbook` → `Close` → `Workbooks.Open path` |
| E02 | `ws.Cells(r, c).Value2` |
| E03 | 地道版：`ws.Range(addr).Value2 = arr`（`Dim arr(1 To n, 1 To m)`）；naive 双重 For |
| E04 | `Application.Calculation = xlCalculationManual` → 写 → `Application.Calculate` → 复位 |
| E05 | `rng.Interior.Color = RGB(...)`（VBA `RGB()` 帮你转 BGR）；`rng.NumberFormat = "0.00"` |
| E06 | `ws.Columns("A:C").AutoFit` |
| E07 | `Worksheets.Add`，`ws.Name = "Data"`，`Application.DisplayAlerts = False` 后 `Delete` |
| E08 | `rng.Replace`, `rng.Sort`, `rng.AutoFilter` —— 具名常量直接可用 |
| E09 | `ws.ChartObjects.Add(...).Chart`，`.ChartType = xlColumnClustered` |
| E10 | `wb.ExportAsFixedFormat xlTypePDF, pdfPath` |
| E11 | **VBA 是宏宿主本身**：直接定义并 `Application.Run "Macro"` 或直接调用过程；最自然 |
| E12 | 若新建了独立实例：`wb.Close False` → `app.Quit` → `Set app = Nothing`；宿主内则无独立进程 |

## 资源释放策略（E12）

- 宿主内（用 ActiveWorkbook）：不产生额外进程，置对象 `= Nothing` 即可。
- 新建独立 `Excel.Application`：`app.Quit` + `Set ... = Nothing`；VBA 无 `ReleaseComObject`，靠置空 + 引擎回收。

## 目录约定（P1）

```
vba/
├── README.md
└── src/
    ├── modExcelTasks.bas   # 导出的标准模块（E01..E12 各 Sub）
    ├── modRunner.bas       # 跑单任务/全套 + Timer 计时
    └── README-import.md    # 如何导入 .bas 到 Excel 并运行（含命令行触发方式）
```

VBA 代码以可导出的 `.bas` 文本形式入库，便于 diff 与版本管理；附导入/运行说明。

## 待办（P1）

- [ ] `modExcelTasks.bas`：E01–E12
- [ ] `modRunner.bas`：入口 + `Timer` 计时
- [ ] 导入/headless 触发说明（如用 `cscript` + xlsm，或命令行 `/m`）
- [ ] native 实测回填能力矩阵
