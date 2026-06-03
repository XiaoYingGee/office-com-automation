# 01 · Excel 对象模型详解

> 承接 [00-com-overview.md](00-com-overview.md)。本文聚焦 Excel 自动化对象树中**做编辑最常用**的对象与成员，
> 语言无关。各语言 README 里的 `E01–E12` 实现都映射到这里的对象。

## 1. 对象树主干

```
Application
├─ Visible / DisplayAlerts / ScreenUpdating / Calculation   ← 全局开关（性能/静默关键）
├─ Workbooks
│   ├─ .Add()              新建工作簿
│   ├─ .Open(path)         打开
│   └─ Item(i) / [name]    取某个 Workbook
│       └─ Workbook
│           ├─ .Save() / .SaveAs(path, fileFormat)
│           ├─ .ExportAsFixedFormat(0, pdfPath)   ← 导出 PDF
│           ├─ .Close(SaveChanges)
│           └─ Worksheets / Sheets
│               ├─ .Add() / Item(i) / [name]
│               └─ Worksheet
│                   ├─ .Name
│                   ├─ .Cells(row, col)            单格（1-based）
│                   ├─ .Range("A1:C3")             区域
│                   ├─ .Rows / .Columns            整行/整列
│                   ├─ .UsedRange                  已用区域
│                   └─ .ChartObjects / .Shapes     图表与图形
```

## 2. Range —— 编辑的主战场

`Range` 代表一个或多个单元格，几乎所有编辑都落在它上面。

### 2.1 取值/赋值

| 成员 | 说明 |
|---|---|
| `Range.Value` | 带类型封送（日期→Date、货币→Currency），较慢 |
| `Range.Value2` | **不做日期/货币特殊封送**，更快、更可预测；批量读写首选 |
| `Range.Text` | 单元格显示文本（只读，含格式） |
| `Range.Formula` | A1 样式公式字符串，如 `=A1+B1` |
| `Range.FormulaR1C1` | R1C1 样式公式，如 `=RC[-1]+RC[-2]` |

**批量读写（核心技巧）**——一次往返搞定整片区域：

```text
写： 把一个 2 维数组赋给 Range("A1:B1000").Value2
读： rows = Range("A1:B1000").Value2   → 得到 2 维数组
```

各语言数组形态不同（PowerShell `object[,]`、Python 元组的元组、C# `object[,]`、JScript SafeArray、
C++ `VARIANT`/`SAFEARRAY`），但语义一致：**1 次封送 vs N 次封送**。

### 2.2 定位与遍历

| 成员 | 用途 |
|---|---|
| `Cells(r, c)` | 按行列号取单格（1-based） |
| `Range(cell1, cell2)` | 两角定义矩形区域 |
| `Range("A1").End(xlDown/xlUp/xlToRight/xlToLeft)` | 跳到数据边界（`xlUp=-4162` 等常量） |
| `UsedRange` | 已使用区域，遍历起点 |
| `Range.Offset(dr, dc)` / `.Resize(rows, cols)` | 相对定位/改尺寸 |

### 2.3 格式化（E05）

| 对象/属性 | 作用 |
|---|---|
| `Range.Font.Bold / .Italic / .Size / .Name / .Color` | 字体 |
| `Range.Interior.Color` | 填充色（**BGR** 整数，如黄色 `65535`） |
| `Range.Borders` | 边框 |
| `Range.NumberFormat` | 数字格式，如 `"0.00"`、`"yyyy-mm-dd"`、`"#,##0"` |
| `Range.HorizontalAlignment / .VerticalAlignment` | 对齐（`xlCenter=-4108` 等） |
| `Range.Merge() / .UnMerge()` | 合并单元格 |

> 颜色是 **BGR**（蓝绿红），不是 RGB。`RGB(r,g,b) = r + g*256 + b*65536`。

### 2.4 结构编辑（E06）

| 操作 | 调用 |
|---|---|
| 插入行/列 | `Rows(i).Insert()` / `Columns(i).Insert()` |
| 删除行/列 | `Rows(i).Delete()` / `Columns(i).Delete()` |
| 自动列宽 | `Columns("A:C").AutoFit()` |
| 设行高/列宽 | `Rows(i).RowHeight = 20` / `Columns(i).ColumnWidth = 12` |

## 3. 公式与重算（E04）

- 写公式：设 `Range.Formula`。
- Excel 默认自动重算；批量写入时常先关自动重算提速，最后再强制算：
  - `Application.Calculation = xlCalculationManual` （`-4135`）
  - 批量写…
  - `Application.Calculate()` 或 `Worksheet.Calculate()` 强制重算
  - 恢复 `xlCalculationAutomatic`（`-4105`）
- 读结果：算完后读 `Range.Value2`。

## 4. 全局性能/静默开关

驱动 Excel 时，**先关掉这些再批量操作**能显著提速并避免弹窗：

| 开关 | 设为 | 效果 |
|---|---|---|
| `Application.ScreenUpdating` | `False` | 不刷新屏幕 |
| `Application.Calculation` | `xlCalculationManual` | 暂停自动重算 |
| `Application.DisplayAlerts` | `False` | 抑制"是否保存"等弹窗 |
| `Application.EnableEvents` | `False` | 关事件，避免触发宏 |
| `Application.Visible` | `False` | 后台运行 |

> 收尾记得**恢复**这些开关（尤其在复用同一个 Application 实例时）。

## 5. 多工作表与跨表引用（E07）

- 新增：`Worksheets.Add(After:=Worksheets(Count))`
- 改名：`Worksheet.Name = "Data"`
- 删除：`DisplayAlerts=False` 后 `Worksheet.Delete()`
- 跨表公式：`='Data'!A1`（含空格/特殊字符的表名要加单引号）

## 6. 图表（E09）与导出（E10）

- 图表：`Worksheet.ChartObjects.Add(left, top, width, height)` → `.Chart.SetSourceData(Range)` →
  `.Chart.ChartType = 51`（柱状图 `xlColumnClustered`）。
- 导出 PDF：`Workbook.ExportAsFixedFormat(Type:=0, Filename:=pdfPath)`（`0 = xlTypePDF`）。

## 7. 常用枚举常量速查

早绑定语言有具名常量；晚绑定（PowerShell/JScript/Python-Dispatch）常需直接写数字：

| 常量 | 值 | 用途 |
|---|---|---|
| `xlUp` | -4162 | `End` 方向 |
| `xlDown` | -4121 | `End` 方向 |
| `xlToRight` | -4161 | `End` 方向 |
| `xlCenter` | -4108 | 居中对齐 |
| `xlCalculationManual` | -4135 | 手动重算 |
| `xlCalculationAutomatic` | -4105 | 自动重算 |
| `xlColumnClustered` | 51 | 簇状柱形图 |
| `xlTypePDF` | 0 | 导出 PDF |
| `xlValues` | -4163 | 查找/选择性粘贴 |
| `xlPasteValues` | -4163 | 选择性粘贴值 |
| `xlOpenXMLWorkbook` | 51 | .xlsx 文件格式（SaveAs） |

> 完整枚举见微软 Excel VBA 参考；本表覆盖 `E01–E12` 所需。

---

← 上一篇 [00-com-overview.md](00-com-overview.md) ｜ 标准任务集 → [../spec/excel-tasks.md](../spec/excel-tasks.md)
