# Excel COM Object Model — AI Quick Reference

> 用于 `excel_execute_code` 工具。代码中可用 `app`, `wb`, `ws` 变量。
> 设置 `result` 变量返回数据给调用者。

---

## 核心对象层级

```
app (Excel.Application)
 └── Workbooks
      └── wb (Workbook)
           ├── Worksheets / Sheets
           │    └── ws (Worksheet)
           │         ├── Range("A1") / Range("A1:C10") / Cells(row, col)
           │         ├── Rows(n) / Columns(n)
           │         ├── UsedRange
           │         ├── ChartObjects()
           │         ├── Shapes
           │         ├── Hyperlinks
           │         ├── PageSetup
           │         ├── PivotTables(name)
           │         └── Names
           ├── Charts
           ├── PivotCaches()
           └── Names
```

---

## 常用操作速查

### 单元格读写

```python
ws.Range("A1").Value2 = "Hello"           # 写值
ws.Range("A1").Value2                      # 读值
ws.Range("A1").Formula = "=SUM(B1:B10)"   # 写公式
ws.Range("A1").Text                        # 读显示文本
ws.Cells(row, col).Value2 = 42            # 按行列号（1-based）
```

### 区域批量操作

```python
# 写二维数组（一次封送，极快）
ws.Range("A1:C3").Value2 = ((1,2,3),(4,5,6),(7,8,9))

# 读二维数组
data = ws.Range("A1:C100").Value2  # → tuple of tuples

# 清除
ws.Range("A1:Z100").ClearContents()  # 只清值
ws.Range("A1:Z100").Clear()          # 清值+格式

# 复制粘贴
ws.Range("A1:A10").Copy()
ws.Range("B1").PasteSpecial(Paste=-4163)  # xlPasteValues
app.CutCopyMode = False

# 自动填充
ws.Range("A1:A2").AutoFill(ws.Range("A1:A100"))
```

### 格式化

```python
rng = ws.Range("A1:E1")
rng.Font.Bold = True
rng.Font.Italic = True
rng.Font.Size = 14
rng.Font.Name = "微软雅黑"
rng.Font.Color = 0x0000FF          # BGR! 红=0x0000FF, 蓝=0xFF0000
rng.Interior.Color = 0x00FFFF      # 背景色（BGR）
rng.NumberFormat = "#,##0.00"       # 数字格式
rng.HorizontalAlignment = -4108    # xlCenter
rng.VerticalAlignment = -4108
rng.WrapText = True

# 边框
rng.Borders.LineStyle = 1          # xlContinuous
rng.Borders.Weight = 2             # xlThin
rng.Borders(7).LineStyle = 1       # 7=左, 8=上, 9=下, 10=右, 11=内竖, 12=内横
```

### 行列操作

```python
ws.Rows(5).Insert()                # 在第5行前插入
ws.Rows("5:10").Delete()           # 删除5-10行
ws.Columns("B").Insert()           # 在B列前插入
ws.Columns("B:D").Delete()

ws.Rows(1).RowHeight = 30          # 设置行高
ws.Columns("A").ColumnWidth = 20   # 设置列宽
ws.UsedRange.Columns.AutoFit()     # 自适应列宽

ws.Rows("2:10").Group()            # 分组
ws.Rows("2:10").Hidden = True      # 隐藏
```

### 工作表管理

```python
ws_new = wb.Worksheets.Add()       # 新增
ws_new.Name = "新表"
ws.Copy(After=wb.Worksheets(wb.Worksheets.Count))  # 复制
ws.Move(Before=wb.Worksheets(1))   # 移动
ws.Delete()                         # 删除（需 DisplayAlerts=False）
ws.Protect(Password="123")          # 保护
ws.Unprotect(Password="123")
```

### 窗口

```python
app.Goto(ws.Cells(3, 2))
app.ActiveWindow.FreezePanes = True   # 冻结
app.ActiveWindow.FreezePanes = False  # 取消冻结
```

### 排序与筛选

```python
# 排序
rng = ws.Range("A1:E100")
rng.Sort(Key1=ws.Range("C1"), Order1=1, Header=1)  # 1=xlAscending, 2=xlDescending

# 自动筛选
ws.Range("A1:E1").AutoFilter(Field=3, Criteria1=">100")
ws.AutoFilterMode = False  # 关闭筛选
```

### 查找替换

```python
ws.Cells.Replace(What="旧值", Replacement="新值")
found = ws.Cells.Find(What="搜索文本")
if found:
    result = found.Address  # → "$C$5"
```

### 数据验证

```python
rng = ws.Range("B2:B100")
rng.Validation.Delete()
rng.Validation.Add(3, 1, 1, "选项A,选项B,选项C")  # 3=xlValidateList
```

### 命名区域

```python
wb.Names.Add("数据区", "=Sheet1!$A$1:$E$100")
wb.Names("数据区").Delete()
# 使用
ws.Range("数据区").Value2
```

### 条件格式

```python
rng = ws.Range("C2:C100")
# 大于 1000 标红
cf = rng.FormatConditions.Add(1, 5, "1000")  # 1=xlCellValue, 5=xlGreater
cf.Interior.Color = 0x0000FF
cf.Font.Bold = True
```

### 图表

```python
co = ws.ChartObjects().Add(100, 100, 400, 300)  # left, top, width, height
chart = co.Chart
chart.ChartType = 51              # xlColumnClustered
chart.SetSourceData(ws.Range("A1:B10"))
chart.HasTitle = True
chart.ChartTitle.Text = "销售趋势"

# 常用类型: 51=柱状, 4=折线, 5=饼图, 57=条形, 1=面积, -4169=散点
```

### 图片

```python
pic = ws.Shapes.AddPicture(r"C:\path\image.png", False, True, 0, 0, 200, 150)
pic.Name = "MyPic"
ws.Shapes("MyPic").Delete()
```

### 批注

```python
ws.Range("A1").AddComment("这是批注")
ws.Range("A1").Comment.Delete()
ws.Range("A1").Comment.Text()      # 读取批注
```

### 超链接

```python
ws.Hyperlinks.Add(Anchor=ws.Range("A1"), Address="https://example.com",
                  TextToDisplay="点击这里")
```

### 导出

```python
wb.ExportAsFixedFormat(0, r"C:\output.pdf")  # 0=xlTypePDF

# 页面设置
ps = ws.PageSetup
ps.Orientation = 2      # xlLandscape (1=Portrait)
ps.PaperSize = 9        # A4
ps.LeftMargin = 36      # points (72pt = 1 inch)
ps.CenterHeader = "&B公司报表"
ps.FitToPagesWide = 1
ps.FitToPagesTall = 0   # 0=auto
```

### 数据透视表

```python
pc = wb.PivotCaches().Create(1, ws.Range("A1:E100"))  # 1=xlDatabase
pt = pc.CreatePivotTable(wb.Sheets("汇总").Range("A1"), "MyPivot")
pt.PivotFields("产品").Orientation = 1    # xlRowField
pt.PivotFields("地区").Orientation = 2    # xlColumnField
pt.AddDataField(pt.PivotFields("金额"))   # 默认求和
pt.RefreshTable()
```

### 宏执行

```python
result = app.Run("MacroName", arg1, arg2)
```

---

## 常量速查

| 常量 | 值 | 用途 |
|------|:---:|------|
| xlAscending | 1 | Sort 升序 |
| xlDescending | 2 | Sort 降序 |
| xlYes (Header) | 1 | 有表头 |
| xlCenter | -4108 | 居中对齐 |
| xlLeft | -4131 | 左对齐 |
| xlRight | -4152 | 右对齐 |
| xlContinuous | 1 | 实线边框 |
| xlThin | 2 | 细线 |
| xlMedium | -4138 | 中线 |
| xlThick | 4 | 粗线 |
| xlPasteValues | -4163 | 粘贴值 |
| xlPasteFormulas | -4123 | 粘贴公式 |
| xlPasteFormats | -4122 | 粘贴格式 |
| xlTypePDF | 0 | 导出 PDF |
| xlColumnClustered | 51 | 柱状图 |
| xlLine | 4 | 折线图 |
| xlPie | 5 | 饼图 |
| xlRowField | 1 | 透视表行字段 |
| xlColumnField | 2 | 透视表列字段 |
| xlDatabase | 1 | 透视表数据源 |

---

## BGR 颜色

COM 使用 **BGR** 而非 RGB！

| 颜色 | BGR 值 | 公式 |
|------|--------|------|
| 红 | `0x0000FF` | R + G×256 + B×65536 |
| 绿 | `0x00FF00` | |
| 蓝 | `0xFF0000` | |
| 黄 | `0x00FFFF` | |
| 白 | `0xFFFFFF` | |
| 黑 | `0x000000` | |

---

## 注意事项

- **索引从 1 开始**（Cells(1,1) = A1）
- **设置 result 变量**返回数据给调用者
- **错误会自动捕获**并返回 traceback，AI 可据此修正
- **app.DisplayAlerts = False** 已设置，不会弹对话框
- **app.ScreenUpdating = False** 可加速大批量操作（完事后设回 True）
