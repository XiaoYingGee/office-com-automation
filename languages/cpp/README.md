# C++ (原生 COM) · 实现 Plan

> 参赛语言之一。最底层、最可控——直接走 OLE Automation（`IDispatch`/`#import`）。
> 实际代码在 P1 写入本目录 `src/`。

## 定位

- **绑定**：两种路线 ——
  1. **`#import`**：从 `EXCEL.EXE` 类型库生成智能包装（`_ApplicationPtr` 等），早绑定式，写法接近其它语言。
  2. **裸 `IDispatch`**：`CLSIDFromProgID` + `CoCreateInstance` + `Invoke`，最底层、样板最多。
- **优势**：无托管运行时、性能上限高、对封送/生命周期完全掌控。
- **劣势**：样板代码多、`VARIANT`/`SAFEARRAY` 手工管理、开发慢。

## 环境前提

- Windows + Excel。
- MSVC（含 ATL/COM 支持）；建议用 ATL `CComPtr`/`CComVariant` 简化引用与释放。
- `CoInitializeEx` / `CoUninitialize` 包裹 COM 使用。
- Wine：产物为原生 PE，依赖 Wine 的 `ole32`/`oleaut32`/automation 覆盖度，🟡 P2 实测。

## 接入方式（#import 路线）

```cpp
#import "C:\\Program Files\\...\\EXCEL.EXE" rename("DialogBox","ExcelDialogBox") \
    rename("RGB","ExcelRGB") rename("CopyFile","ExcelCopyFile") no_dual_interfaces

#include <atlbase.h>

CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);
{
    Excel::_ApplicationPtr app;
    app.CreateInstance(__uuidof(Excel::Application));
    app->Visible = VARIANT_FALSE;
    app->DisplayAlerts = VARIANT_FALSE;
    Excel::_WorkbookPtr wb = app->Workbooks->Add();
    Excel::_WorksheetPtr ws = wb->Worksheets->GetItem(1);
}
CoUninitialize();
```

## E01–E12 实现要点

| 任务 | 关键写法 / 坑 |
|---|---|
| E01 | `app->Workbooks->Add()`；`wb->SaveAs(path, 51)`；`_bstr_t` 转字符串 |
| E02 | `ws->Cells->Item[r][c]->Value2`；用 `_variant_t` 收发 |
| E03 | **核心**：构造 2D `SAFEARRAY`（`VT_VARIANT`）一次性赋 `Range->Value2`；naive 用双重循环 `Item`。C++ 在这里能拿到性能优势，但 SAFEARRAY 手工管理是难点 |
| E04 | `app->Calculation = xlCalculationManual` → 写 → `app->Calculate()` → 复位 |
| E05 | `rng->Interior->Color`（BGR `long`）；`rng->NumberFormat = L"0.00"` |
| E06 | `ws->Columns->Item["A:C"]->AutoFit()` |
| E07 | `wb->Worksheets->Add()`；`ws->Name = L"Data"` |
| E08 | `rng->Replace(...)`/`Sort`/`AutoFilter`；缺省参用 `vtMissing` |
| E09 | `ws->ChartObjects()->Add(...)->Chart` |
| E10 | `wb->ExportAsFixedFormat(Excel::xlTypePDF, path)` |
| E11 | `app->Run("MacroName")` |
| E12 | **重点**：用 `CComPtr`/智能 `_xxxPtr` 让引用随作用域析构；裸接口须手工 `Release()`；`app->Quit()` 后确保所有 Ptr 出作用域，再 `CoUninitialize()` |

## SAFEARRAY 难点（E03 专题）

地道的批量写需要构造 `VT_ARRAY|VT_VARIANT` 的二维 `SAFEARRAY`：

```text
SafeArrayCreate(VT_VARIANT, 2, bounds[rows, cols])
→ 逐元素 SafeArrayPutElement（构造期）
→ 包进 VARIANT 赋给 Range->Value2（1 次封送）
→ SafeArrayDestroy / 由 _variant_t 析构
```

这是 C++ 样板最重、也最体现"可控 vs 啰嗦"权衡的地方，PK D3 重点观察。

## 资源释放策略（E12）

- 优先 `#import` + 智能指针（`_xxxPtr` 引用计数自动），作用域结束自动 `Release`。
- 裸 `IDispatch` 路线：每个 `Invoke` 拿到的接口指针都要 `Release`；用 `CComPtr` 包裹避免漏放。
- 顺序：业务对象出作用域 → `app->Quit()` → app Ptr 出作用域 → `CoUninitialize()` → `tasklist` 校验无残留。

## 目录约定（P1）

```
cpp/
├── README.md
└── src/
    ├── CMakeLists.txt / *.vcxproj
    ├── excel_com.h/.cpp   # RAII 封装：CoInit 守卫 + app 启停 + SAFEARRAY 写 helper
    ├── tasks/e01..e12.cpp
    └── main.cpp           # CLI：跑单任务/全套 + QueryPerformanceCounter 计时
```

## 待办（P1）

- [ ] 选定 `#import` 路线 + ATL 智能指针骨架
- [ ] `excel_com`：CoInit RAII 守卫 + SAFEARRAY 批量写 helper
- [ ] E01–E12
- [ ] `main.cpp` CLI + 高精度计时
- [ ] native 实测回填能力矩阵（记录 SAFEARRAY 方案与性能）
