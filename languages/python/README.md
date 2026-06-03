# Python (pywin32) · 实现 Plan

> 参赛语言之一。工程化友好、生态丰富，定为 PK 的**参照实现**。实际代码在 P1 写入本目录 `src/`。

## 定位

- **绑定**：早绑定 `win32com.client.gencache.EnsureDispatch`（带常量/类型缓存）或晚绑定 `Dispatch`。
- **优势**：可读性好、异常处理强、易工程化与测试；`constants.xlUp` 具名常量。
- **劣势**：依赖 pywin32（PE 扩展），**Wine 下兼容性偏谨慎**。

## 环境前提

- Windows + Excel。
- Python 3.x + `pywin32`：`pip install pywin32`。
- 早绑定首次需生成类型缓存（`EnsureDispatch` 自动；或 `makepy`）。
- Wine：需 Windows 版 Python + pywin32 在 prefix 内运行，风险较高，P2 实测。

## 接入方式

```python
import win32com.client as win32
from win32com.client import constants as c

app = win32.gencache.EnsureDispatch('Excel.Application')  # 早绑定
app.Visible = False
app.DisplayAlerts = False
```

## E01–E12 实现要点

| 任务 | 关键写法 / 坑 |
|---|---|
| E01 | `wb.SaveAs(path, FileFormat=51)`；用原始字符串路径 `r'...'` |
| E02 | `ws.Cells(r, c).Value2`（1-based） |
| E03 | 地道版：`ws.Range(addr).Value2 = tuple(tuple(row) for row in data)`；naive 版双重 for |
| E04 | `app.Calculation = c.xlCalculationManual` → 写 → `app.Calculate()` → 复位 |
| E05 | `rng.Interior.Color`（BGR int）；`rng.NumberFormat = '0.00'` |
| E06 | `ws.Columns('A:C').AutoFit()` |
| E07 | `wb.Worksheets.Add()`；`ws.Name='Data'` |
| E08 | `rng.Replace()`, `rng.Sort()`, `rng.AutoFilter()`；可选参用关键字传 |
| E09 | `ws.ChartObjects().Add(l,t,w,h).Chart` 链式 |
| E10 | `wb.ExportAsFixedFormat(0, pdf_path)` |
| E11 | `app.Run('MacroName')`；注入 VBA 需信任设置 |
| E12 | 让子对象引用离开作用域 + `wb.Close(False)` + `app.Quit()`；必要时 `del` + `gc.collect()` |

## 资源释放策略（E12）

Python 引用计数通常能回收，但持有的 COM 子对象需显式断引：

```python
try:
    ...
finally:
    wb.Close(False)
    app.Quit()
    del rng, ws, wb, app
    import gc; gc.collect()
```
注意早绑定生成的 `gen_py` 缓存目录（已在 .gitignore）。

## 目录约定（P1）

```
python/
├── README.md
└── src/
    ├── excel_com.py    # 封装：app 上下文管理器(with) + 批量写 helper
    ├── tasks/e01..e12.py
    └── run.py          # CLI：argparse 跑单任务/全套 + 计时
```

建议把 App 封成上下文管理器，`__exit__` 里统一释放，从根上治理 E12。

## 待办（P1）

- [ ] `excel_com.py`：`with ExcelApp() as app` 上下文 + 安全释放 + 批量封送
- [ ] E01–E12
- [ ] `run.py` 入口 + 计时
- [ ] native 实测回填能力矩阵
