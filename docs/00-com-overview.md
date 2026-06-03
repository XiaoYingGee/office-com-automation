# 00 · COM 概述：Office 自动化的底层机制

> 目标读者：准备用任意语言驱动 Office 的工程师。本文是**语言无关**的基础，后续各语言 README 默认你已读过本篇。

## 1. COM 是什么

**COM（Component Object Model）** 是 Windows 上的二进制级组件互操作标准。一个组件把功能封装在
**接口（interface）**后面，调用方只依赖接口契约，不关心实现语言。Office 把它的功能以一组 COM 服务器
的形式暴露出来：

| 程序标识（ProgID） | 类标识（CLSID 概念） | 暴露的根对象 |
|---|---|---|
| `Excel.Application` | Excel 自动化服务器 | `Application` |
| `Word.Application` | Word 自动化服务器 | `Application` |
| `PowerPoint.Application` | PowerPoint 自动化服务器 | `Application` |

当你"New 一个 `Excel.Application`"时，Windows 通过注册表用 ProgID 找到 CLSID，启动（或复用）
`EXCEL.EXE` 进程，并把它的根 `Application` 接口指针交给你。**你随后做的一切编辑，都是隔着进程边界
在调这个 COM 服务器的方法。**

## 2. 对象模型：一棵可导航的树

Office 自动化对象模型是一棵对象树。"编辑文档"= 从根对象导航到目标对象 + 设属性/调方法。Excel 的主干：

```
Application                 Excel 进程
 └─ Workbooks               打开的工作簿集合
     └─ Workbook            单个 .xlsx
         └─ Worksheets      工作表集合
             └─ Worksheet
                 └─ Range   单元格/区域 —— 编辑主战场
                     ├─ Value / Value2
                     ├─ Formula / FormulaR1C1
                     ├─ Font / Interior / Borders
                     └─ NumberFormat
```

详见 [01-excel-object-model.md](01-excel-object-model.md)。

## 3. 早绑定 vs 晚绑定

调用 COM 方法有两种解析方式，深刻影响开发体验、性能和 Wine 兼容性：

| | 晚绑定（late binding） | 早绑定（early binding） |
|---|---|---|
| 解析时机 | 运行时按名字через `IDispatch::GetIDsOfNames` | 编译期按类型库（TLB）解析 |
| 代表 | PowerShell `New-Object -ComObject`、JScript `CreateObject`、Python `Dispatch` | C# Interop、C++ `#import`、Python `EnsureDispatch` |
| 优点 | 无需类型库、写起来直接 | 有智能感知/常量/类型检查、调用更快 |
| 缺点 | 无编译检查、每次调用多一次名字解析开销 | 需要 TLB、版本耦合 |
| 常量（如 `xlUp`） | 需自己写数字 | 可用具名枚举 |

**经验**：批量调用密集的场景早绑定更快；快速脚本晚绑定更省事。

## 4. 跨进程封送（marshaling）与"批量优于循环"

你的代码和 `EXCEL.EXE` 是**两个进程**。每一次属性读写/方法调用都要把参数**封送**过进程边界——
这是 COM 自动化最大的性能税。

```
逐格写 10,000 个单元格  ≈ 10,000 次跨进程往返   → 数秒~数十秒
一次性把 10,000 个值作为二维数组赋给 Range.Value2 → 1 次往返 → 毫秒级
```

> **第一性能原则：能批量就别循环。** 读写大区域时，用二维数组一次性 set/get `Range.Value2`。
> 这条规则对所有语言成立，也是 `E03` 基准任务存在的意义。

## 5. 引用计数与"僵尸进程"

COM 用**引用计数**管理对象生命周期（`AddRef`/`Release`）。你每拿到一个子对象（`Workbook`、
`Worksheet`、`Range`…）都持有一份引用。**只要还有一份引用没释放，`EXCEL.EXE` 就不会退出**——
即使你以为已经 `Quit()` 了，进程仍残留在内存里（俗称僵尸 EXCEL.EXE）。

各语言的释放责任：

| 语言 | 释放方式 |
|---|---|
| PowerShell | `[Runtime.InteropServices.Marshal]::ReleaseComObject($obj)` + `$obj=$null` |
| Python(pywin32) | 让引用离开作用域 / `del`，必要时 `gc.collect()`；`app.Quit()` |
| C# | `Marshal.ReleaseComObject(obj)` 自底向上，`GC.Collect()` + `WaitForPendingFinalizers()` |
| C++ | `pObj->Release()` 或用 `CComPtr` 智能指针自动管理 |
| JScript/VBA | 引擎回收 + 显式 `Quit`；置 `null`/`Nothing` |

> `E12`（资源释放 & 僵尸进程清理）是一个独立的**工程正确性**评分项，专门考察这点。

## 6. 各语言接入 COM 的通用机制（速查）

```powershell
# PowerShell（晚绑定）
$app = New-Object -ComObject Excel.Application
```
```python
# Python / pywin32
import win32com.client as win32
app = win32.gencache.EnsureDispatch('Excel.Application')   # 早绑定
app = win32.Dispatch('Excel.Application')                  # 晚绑定
```
```csharp
// C# / Interop（需引用 Microsoft.Office.Interop.Excel）
var app = new Microsoft.Office.Interop.Excel.Application();
```
```javascript
// JScript (WSH)
var app = WScript.CreateObject("Excel.Application");
```
```cpp
// C++（#import 生成包装，或裸 IDispatch）
// #import "EXCEL.EXE" 后用生成的智能指针；或 CLSIDFromProgID + CoCreateInstance
```

## 7. 与"非 COM"方案的边界

| 方案 | 需装 Office | 跑宏/重算公式 | 适用 |
|---|---|---|---|
| **COM 自动化（本项目）** | 是 | 能 | 需要 Excel 真实引擎、重算、宏、PDF 导出 |
| openpyxl / EPPlus / OpenXML SDK | 否 | 否（只写公式字符串） | 纯数据读写、服务器批处理 |

本项目刻意研究 **COM**——即调用 Office **自身的编辑/计算引擎**，而非改 XML。

---

下一篇 → [01-excel-object-model.md](01-excel-object-model.md)
