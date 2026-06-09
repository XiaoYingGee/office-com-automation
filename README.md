# Office COM Automation —— Excel 自动化

> 用 **COM** 调用 Microsoft Office Excel 的原生编辑接口，
> 四种方式驱动 Excel 完成自动化任务，验证 IPC 边界是性能瓶颈而非语言本身。

[![status](https://img.shields.io/badge/status-implemented-brightgreen)](#) [![license](https://img.shields.io/badge/license-MIT-blue)](LICENSE)

---

## 这是什么

COM（Component Object Model）是 Windows 上自动化 Office 的核心机制。Office 把 `Excel.Application`
暴露为 COM 服务器，支持 COM 的语言都能驱动它的**真实引擎**去读写、编辑、计算、导出文档。

本项目提供一套 Excel 自动化工具，支持四种后端（2×2 矩阵）：

```
                  Out-of-Process (IPC)        In-Process (零 IPC)
                  ─────────────────────────   ─────────────────────────
Python            pywin32 (baseline)          Python COM Add-in
VBA/C#            —                           VBA / C# COM Add-in
```

| 方式 | 接入方式 | 性能 | 备注 |
|------|----------|------|------|
| **pywin32** | Python → 跨进程 COM IPC | 基线 | 每属性一次 IPC 往返 |
| **VBA** | Python → `Application.Run` | ~10-50x 快 | 零 IPC，进程内 |
| **Python Add-in** | COMAddIn.Object → 进程内 Python | ≈ VBA | 同一 Python 代码，移入进程内 |
| **C# Add-in** | COMAddIn.Object → 进程内 .NET | ≈ VBA | 早绑定 vtable 调用 |

## Benchmark 结论

VBA 和 in-process add-in 在 IPC 密集型操作上比 pywin32 快 **10–50 倍**。
瓶颈是跨进程 COM 边界，不是语言。详见 [benchmark.md](benchmark.md)。

## 目录导航

```
docs/           COM 概述、Excel 对象模型
skills/         Excel editor skill（pywin32 + VBA 后端、benchmark 脚本）
python_addin/   Python in-process COM add-in
csharp_addin/   C# in-process COM add-in (.NET Framework 4.8)
```

## 快速开始

```bash
# 基本用法（pywin32）
python skills/excel-editor/scripts/excel_editor.py test.xlsx --create --inspect

# VBA 后端
python skills/excel-editor/scripts/excel_editor.py test.xlsx --inspect --backend vba

# 注册 Python add-in（然后重启 Excel）
python python_addin/excel_pyaddin.py

# 构建并注册 C# add-in
cd csharp_addin/ExcelEditorAddin && dotnet build -c Release && cd ..
powershell ./register.ps1

# 跑 benchmark（4 后端对比）
python skills/excel-editor/scripts/benchmark.py --all --size empty,medium
```

## License

[MIT](LICENSE)
