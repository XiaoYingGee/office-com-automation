# Office COM Automation —— Excel 自动化

> 用 **COM** 调用 Microsoft Office Excel 的原生编辑接口，
> 通过 Python (pywin32) + VBA 两种方式驱动 Excel 完成自动化任务。

[![status](https://img.shields.io/badge/status-implemented-brightgreen)](#) [![license](https://img.shields.io/badge/license-MIT-blue)](LICENSE)

---

## 这是什么

COM（Component Object Model）是 Windows 上自动化 Office 的核心机制。Office 把 `Excel.Application`
暴露为 COM 服务器，支持 COM 的语言都能驱动它的**真实引擎**去读写、编辑、计算、导出文档。

本项目提供一套 Excel 自动化工具（excel-editor skill），支持两种后端：

| 方式 | 接入方式 | 会话模型 | 备注 |
|------|----------|----------|------|
| **pywin32** | Python → 跨进程 COM IPC | 持久会话，打开一次多次操作 | 每属性一次 IPC 往返 |
| **VBA** | Python → `Application.Run` → 进程内 VBA | 持久会话，打开一次多次操作 | 零 IPC，最快 |

## Benchmark 结论

VBA 在 IPC 密集型操作上比 pywin32 快 **10–50 倍**（零 IPC vs 逐属性跨进程往返）。
详见 [benchmark.md](benchmark.md)。

## 目录导航

```
docs/      COM 概述、Excel 对象模型
skills/    Excel editor skill（pywin32 + VBA 后端、benchmark 脚本）
```

核心文档：
- [docs/00-com-overview.md](docs/00-com-overview.md) — COM 与对象模型基础
- [docs/01-excel-object-model.md](docs/01-excel-object-model.md) — Excel 对象树详解
- [benchmark.md](benchmark.md) — pywin32 vs VBA 性能对比

## License

[MIT](LICENSE)
