# JScript (WSH) · 实现 Plan

> 参赛语言之一。Windows Script Host 自带 JScript 引擎，**免安装运行时**、纯脚本驱动 COM。
> 实际代码在 P1 写入本目录 `src/`。

## 定位

- **绑定**：晚绑定（`WScript.CreateObject` / `new ActiveXObject`）。
- **优势**：Windows 自带 `cscript`/`wscript`，零安装；轻量、易在脚本化/CI 触发。
- **劣势**：晚绑定无类型/常量；**SafeArray（二维数组）构造繁琐**，影响 E03；语言能力旧。

## 环境前提

- Windows + Excel + WSH（系统自带）。
- 用 `cscript //nologo script.js` 运行（控制台输出友好）。
- Wine：`cscript`/`wscript` 由 Wine 内置提供，可用性不确定（🟡），P2 实测。

## 接入方式

```javascript
// 用 cscript 运行
var app = WScript.CreateObject("Excel.Application");
app.Visible = false;
app.DisplayAlerts = false;
var wb = app.Workbooks.Add();
var ws = wb.Worksheets(1);
```

常量需自定义：`var xl = { calcManual: -4135, calcAuto: -4105, pdf: 0, colClustered: 51 };`

## E01–E12 实现要点

| 任务 | 关键写法 / 坑 |
|---|---|
| E01 | `wb.SaveAs(path, 51)`；路径用双反斜杠或正斜杠 |
| E02 | `ws.Cells(r, c).Value2 = v`；读 `ws.Range("A1").Value2` |
| E03 | **难点**：JScript 原生数组 ≠ COM SafeArray。地道版需借 `Dictionary`/`VBArray` 或 `Scripting` 构造二维数组，或退化为按行写；naive 双重 for。**此项最能拉开与强类型语言的差距** |
| E04 | `app.Calculation = xl.calcManual` → 写 → `app.Calculate()` → 复位 |
| E05 | `rng.Interior.Color = 65535`（BGR）；`rng.NumberFormat = "0.00"` |
| E06 | `ws.Columns("A:C").AutoFit()` |
| E07 | `wb.Worksheets.Add()`；`ws.Name = "Data"` |
| E08 | `rng.Replace(...)`/`rng.Sort(...)`/`rng.AutoFilter(...)`；可选参传 `null` 可能报错，需注意 WSH 对缺省参的处理 |
| E09 | `ws.ChartObjects().Add(l,t,w,h).Chart` |
| E10 | `wb.ExportAsFixedFormat(0, pdfPath)` |
| E11 | `app.Run("MacroName")` |
| E12 | 无 `ReleaseComObject`；`wb.Close(false)` + `app.Quit()` + 置 `null`，依赖引擎回收，**实测是否残留进程** |

## SafeArray 难点（E03 专题）

JScript 没有原生二维 SafeArray 构造能力，把"批量优于循环"落地较麻烦。备选：

1. 按**单行**一次性写（`Range("A1:J1").Value2 = oneDimSafeArray`），循环行 —— 介于地道与 naive 之间。
2. 借助 `new ActiveXObject("System.Collections.ArrayList")` 或 .NET 互操作构造数组（增加依赖）。
3. 退化为逐格（即 naive），作为可用性下限。

> 把这条坑量化记录，正是 PK D2/D3 的看点之一。

## 资源释放策略（E12）

```javascript
wb.Close(false);
app.Quit();
ws = null; wb = null; app = null;
// WSH 无显式 Release，结束后用 tasklist 校验是否残留 EXCEL.EXE
```

## 目录约定（P1）

```
jscript/
├── README.md
└── src/
    ├── excel_com.js   # 公共封装（启停 + 常量表 + 写区域 helper）
    ├── tasks/e01..e12.js
    └── run.js         # cscript 入口：跑单任务/全套 + 计时（用 Date 或 WScript 计时）
```

## 待办（P1）

- [ ] `excel_com.js`：启停 + 常量 + SafeArray 写入折中方案
- [ ] E01–E12
- [ ] `run.js` 入口 + 计时
- [ ] native 实测回填能力矩阵（重点记录 E03 SafeArray 方案与耗时）
