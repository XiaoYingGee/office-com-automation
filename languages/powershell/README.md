# PowerShell · 实现 Plan

> 参赛语言之一。Windows 原生、无需额外运行时，**最快验证 COM** 的一档。本文件是该语言的实现计划，
> 实际代码在 P1 阶段写入本目录 `src/`。

## 定位

- **绑定**：晚绑定（`New-Object -ComObject`）。
- **优势**：Windows 自带，零安装；交互式探索极快。
- **劣势**：晚绑定无类型检查、常量需写数字；大数组性能受脚本引擎影响。

## 环境前提

- Windows + 已安装 Excel（COM 前提）。
- Windows PowerShell 5.1 或 PowerShell 7+。
- Wine 运行时下：PS 在 Wine 历来不稳，PS7 自带运行时更可控 —— 见 [../../docs/wine-sandbox-runtime.md](../../docs/wine-sandbox-runtime.md)，P2 实测。

## 接入方式

```powershell
$app = New-Object -ComObject Excel.Application
$app.Visible = $false
$app.DisplayAlerts = $false
```

常量晚绑定下直接用数字（见 [对象模型常量表](../../docs/01-excel-object-model.md#7-常用枚举常量速查)），
或自定义哈希表集中管理：`$xl = @{ Up=-4162; CalcManual=-4135; PDF=0 }`。

## E01–E12 实现要点

| 任务 | 关键写法 / 坑 |
|---|---|
| E01 | `SaveAs($path, 51)` 指定 xlsx；路径用绝对路径 |
| E02 | `$ws.Cells.Item($r,$c).Value2`；`.Item()` 显式更稳 |
| E03 | 用 `New-Object 'object[,]' $rows,$cols` 构二维数组一次性赋 `Range.Value2`；naive 版用双重 `for` |
| E04 | 设 `$app.Calculation=-4135` → 写公式 → `$app.Calculate()` → 复位 -4105 |
| E05 | `Interior.Color` 用 BGR 整数；`NumberFormat="0.00"` |
| E06 | `$ws.Columns.Item("A:C").AutoFit()` |
| E07 | `$wb.Worksheets.Add()`；删除前 `DisplayAlerts=$false` |
| E08 | `Range.Replace()` / `.Sort()` / `.AutoFilter()` 参数较多，注意可选参用 `[Type]::Missing` |
| E09 | `ChartObjects().Add()` 返回对象链较长，逐级持引用 |
| E10 | `$wb.ExportAsFixedFormat(0, $pdfPath)` |
| E11 | `$app.Run("MacroName")`；注入 VBA 需信任设置 |
| E12 | **重点**：`[Runtime.InteropServices.Marshal]::ReleaseComObject($obj)` 自底向上 + `$obj=$null` + `[GC]::Collect()` |

## 资源释放策略（E12）

PowerShell 最容易留僵尸 `EXCEL.EXE`。约定：

```powershell
$wb.Close($false)
$app.Quit()
foreach ($o in @($range,$ws,$wb,$app)) {
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($o)
}
$range=$ws=$wb=$app=$null
[GC]::Collect(); [GC]::WaitForPendingFinalizers()
```
收尾用 `Get-Process EXCEL -ErrorAction SilentlyContinue` 校验无残留。

## 目录约定（P1）

```
powershell/
├── README.md          # 本文件
└── src/               # P1 实现
    ├── Excel.Com.psm1 # 公共封装（启动/释放/批量写）
    ├── tasks/E01..E12.ps1
    └── run.ps1        # CLI：跑单任务或全套，输出计时
```

## 待办（P1）

- [ ] 公共模块：App 启停 + 安全释放 + 批量封送 helper
- [ ] E01–E12 各脚本
- [ ] `run.ps1` 统一入口 + 计时输出（对接 [benchmark-spec](../../docs/benchmark-spec.md)）
- [ ] native 实测回填 [capability-matrix](../../spec/capability-matrix.md)
