# C# (.NET Interop) · 实现 Plan

> 参赛语言之一。官方 Interop 程序集、早绑定、类型安全、性能好。实际代码在 P1 写入本目录 `src/`。

## 定位

- **绑定**：早绑定，引用 `Microsoft.Office.Interop.Excel`（或 COM Reference / NuGet）。
- **优势**：智能感知、具名枚举（`XlCalculation.xlCalculationManual`）、调用快、可单测。
- **劣势**：Interop + .NET 运行时双重依赖；**Wine 下风险偏高**（运行时 + Interop）。

## 环境前提

- Windows + Excel。
- .NET（建议 .NET 6/8 或 .NET Framework 4.x）+ Excel Interop 程序集（COM 引用或
  `Microsoft.Office.Interop.Excel` NuGet）。
- 设 `[STAThread]`/合适线程模型；启用 `EmbedInteropTypes`（PIA 嵌入，免分发 PIA）。
- Wine：.NET 运行时（wine-mono 或 .NET Core PE）+ Interop，P2 重点实测。

## 接入方式

```csharp
using Excel = Microsoft.Office.Interop.Excel;

var app = new Excel.Application { Visible = false, DisplayAlerts = false };
Excel.Workbook wb = app.Workbooks.Add();
Excel.Worksheet ws = (Excel.Worksheet)wb.Worksheets[1];
```

## E01–E12 实现要点

| 任务 | 关键写法 / 坑 |
|---|---|
| E01 | `wb.SaveAs(path, Excel.XlFileFormat.xlOpenXMLWorkbook)` |
| E02 | `(ws.Cells[r, c] as Excel.Range).Value2` |
| E03 | 地道版：`object[,] data; ws.get_Range(a, b).Value2 = data;` 一次性赋；naive 双重 for |
| E04 | `app.Calculation = Excel.XlCalculation.xlCalculationManual` → 写 → `app.Calculate()` → 复位 |
| E05 | `rng.Interior.Color = ColorTranslator/BGR`；`rng.NumberFormat = "0.00"` |
| E06 | `ws.Columns["A:C"].AutoFit()` |
| E07 | `wb.Worksheets.Add()`；强转 `Excel.Worksheet` |
| E08 | `rng.Replace(...)`, `rng.Sort(...)`, `rng.AutoFilter(...)`；可选参用 `Type.Missing` 或具名参数 |
| E09 | `ws.ChartObjects().Add(...).Chart`，`ChartType = Excel.XlChartType.xlColumnClustered` |
| E10 | `wb.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, pdfPath)` |
| E11 | `app.Run("MacroName")` |
| E12 | **重点**：`Marshal.ReleaseComObject` 自底向上 + `GC.Collect()` + `GC.WaitForPendingFinalizers()` |

## 资源释放策略（E12）

C# 经典坑："两点表达式"会产生未命名的中间 COM 对象而泄漏。约定：

- **避免 `app.Workbooks.Add()` 链式**中产生不可释放的中间引用；必要时把每级对象命名持有再逐个释放。
- 统一封装 `ComCleanup`：

```csharp
wb.Close(false);
app.Quit();
foreach (var o in new object[]{ rng, ws, wb, app })
    if (o != null) Marshal.ReleaseComObject(o);
GC.Collect(); GC.WaitForPendingFinalizers();
```

## 目录约定（P1）

```
csharp/
├── README.md
└── src/
    ├── OfficeComExcel.csproj
    ├── ExcelApp.cs      # IDisposable 封装：启停 + 安全释放 + 批量写
    ├── Tasks/E01..E12.cs
    └── Program.cs       # CLI：跑单任务/全套 + Stopwatch 计时
```

把 `ExcelApp` 实现 `IDisposable`，`using` 作用域结束自动释放，根治 E12。

## 待办（P1）

- [ ] `.csproj` + Interop 引用（`EmbedInteropTypes=true`）
- [ ] `ExcelApp : IDisposable` + 批量封送 helper
- [ ] E01–E12
- [ ] `Program.cs` CLI + Stopwatch
- [ ] native 实测回填能力矩阵
