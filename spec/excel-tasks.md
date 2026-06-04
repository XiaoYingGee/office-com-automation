# Excel 标准任务集 · E01–E12

> **本项目的主心骨。** 这是一份**语言无关的验收口径**：三种语言都实现这同一组任务，才能横向 PK。
> 每个任务给出：目标、必须覆盖的操作点、验收标准、对应的对象模型成员、PK 关联。
>
> 对象模型成员详见 [../docs/01-excel-object-model.md](../docs/01-excel-object-model.md)。

## 约定

- 所有任务在**后台**运行（`Application.Visible = False`），跑批量任务时设静默/性能开关并在收尾恢复。
- 输入数据统一来自 [`../assets/sample-data`](../assets/sample-data)，可复现。
- 产出文件写到各语言 `src/out/`（git 忽略）。
- "验收标准"应能被**独立校验**（用 host 侧 openpyxl 解 .xlsx，或重新打开读值）。

---

## E01 · 工作簿生命周期
**目标**：创建新工作簿、保存为 `.xlsx`、关闭；再打开已存在的 `.xlsx`。
**操作点**：`Workbooks.Add` → `SaveAs(path, xlOpenXMLWorkbook=51)` → `Close` → `Workbooks.Open(path)`。
**验收**：目标路径生成合法 `.xlsx`，可被重新打开；无报错弹窗。
**PK**：D1 基础项；D4（保存格式参数在各语言写法差异）。

## E02 · 单元格读写
**目标**：写入若干单格（字符串/数字/日期/布尔），再读回校验。
**操作点**：`Cells(r,c).Value2`、`Range("B2").Value2`；读 `Value2` 比对。
**验收**：写入值与读回值一致；类型按预期（注意 `Value` vs `Value2` 的日期差异）。
**PK**：D1 基础项。

## E03 · 批量区域写入（★ 性能主项）
**目标**：把 `N×M` 二维数据**一次性**写入区域；并提供 `naive` 逐格版本作对照。
**操作点**：构造二维数组 → `Range(addr).Value2 = array`（地道版）；`for` 循环 `Cells(r,c)=v`（naive 版）。
**验收**：两版写入结果一致；地道版应显著快于 naive 版（量化见 [benchmark-spec](../docs/benchmark-spec.md)）。
**PK**：D2 性能主项（`B1` / `B1-naive`）；体现"批量优于循环"。

## E04 · 公式与重算
**目标**：写入一列公式（如 `=A1*B1`），强制重算，读回计算结果。
**操作点**：`Range.Formula`；`Calculation=xlCalculationManual(-4135)` → 批量写 → `Application.Calculate()` →
读 `Value2` → 恢复 `xlCalculationAutomatic(-4105)`。
**验收**：读回结果等于公式应得值。
**PK**：D1；D2 次要基准（`B2`）。

## E05 · 单元格格式
**目标**：对区域设字体、填充色、数字格式、对齐、合并。
**操作点**：`Font.Bold/Size/Color`、`Interior.Color`（BGR）、`NumberFormat`、`HorizontalAlignment(xlCenter=-4108)`、
`Merge()`。
**验收**：重开文件后格式保留；颜色/数字格式正确（注意 BGR vs RGB）。
**PK**：D1 富格式项；D3（各语言设属性的样板量差异）。

## E06 · 行列结构编辑
**目标**：插入/删除行列，自动列宽，设行高列宽。
**操作点**：`Rows(i).Insert/Delete`、`Columns(i).Insert/Delete`、`Columns("A:C").AutoFit()`、`RowHeight/ColumnWidth`。
**验收**：行列数与内容位移符合预期；AutoFit 后列宽随内容变化。
**PK**：D1。

## E07 · 多工作表
**目标**：新增多个工作表、改名、删除其一、写跨表引用公式。
**操作点**：`Worksheets.Add`、`Worksheet.Name`、`DisplayAlerts=False`+`Delete()`、公式 `='Data'!A1`。
**验收**：表数量/名称正确；跨表公式取到目标表值。
**PK**：D1。

## E08 · 数据操作：查找替换 / 排序 / 筛选
**目标**：在区域内查找替换、按列排序、应用自动筛选。
**操作点**：`Range.Replace(what,replacement)`、`Range.Sort(key, order)`、`Range.AutoFilter(field, criteria)`。
**验收**：替换计数正确；排序后顺序正确；筛选后可见行符合条件。
**PK**：D1（部分语言/Wine 下可能 ⚠️）。

## E09 · 图表生成
**目标**：基于一片数据区域生成簇状柱形图。
**操作点**：`ChartObjects.Add(l,t,w,h)` → `Chart.SetSourceData(Range)` → `Chart.ChartType=xlColumnClustered(51)`。
**验收**：重开文件后存在图表，数据源正确。
**PK**：D1（图形对象，Wine 下重点观察）。

## E10 · 导出 PDF
**目标**：把工作簿导出为 PDF。
**操作点**：`Workbook.ExportAsFixedFormat(Type:=xlTypePDF=0, Filename:=pdfPath)`。
**验收**：生成非空、可打开的 PDF；页面含数据。
**PK**：D1；D2 基准 `B3`；**Wine 下高风险项**（依赖渲染/字体/导出组件）。

## E11 · 运行 / 注入 VBA 宏
**目标**：在工作簿中注入或调用一个 VBA 宏并执行（如批量填色的宏）。
**操作点**：`Application.Run("MacroName", args...)`；或通过 `VBProject` 注入代码（需信任 VBA 对象模型）。
**验收**：宏执行产生预期副作用（单元格被宏修改）。
**PK**：D1 高级项；语言支持度分化大（VBA 原生；外部语言需信任设置）；**Wine 下需实测**。
**注**：注入 VBA 受"信任对 VBA 工程对象模型的访问"安全设置限制，记录各语言/运行时下的可行性。

## E12 · 资源释放 & 僵尸进程清理（★ 工程正确性）
**目标**：任务结束后干净退出，**不残留 `EXCEL.EXE`**。
**操作点**：自底向上释放 COM 引用、`Workbook.Close(False)`、`Application.Quit()`、按语言做
`ReleaseComObject`/`del`/`GC.Collect`/`Release()`。
**验收**：任务结束后进程列表中**无残留** Excel 进程（可用 `tasklist`/`Get-Process` 校验）。
**PK**：独立工程正确性项；与 D3 错误处理联动；**Wine 下进程退出行为需对照原生**。

---

## 任务-维度映射速查

| 任务 | D1 覆盖 | D2 性能 | D3 简洁 | D4/Wine 关注 |
|---|:---:|:---:|:---:|:---:|
| E01 | ● | | | |
| E02 | ● | | | |
| E03 | ● | ★主项 | ● | |
| E04 | ● | ● | | |
| E05 | ● | | ● | |
| E06 | ● | | | |
| E07 | ● | | | |
| E08 | ● | | | ⚠️ |
| E09 | ● | | | ⚠️ |
| E10 | ● | ● | | ⚠️高风险 |
| E11 | ● | | | ⚠️实测 |
| E12 | ● | | ● | ⚠️对照 |

填写实现进度 → [capability-matrix.md](capability-matrix.md)。
