# Office COM Automation —— Excel 自动化

> 用 **COM** 调用 Microsoft Office Excel 的原生编辑接口，
> 通过 Python (pywin32) + VBA 两种方式驱动 Excel 完成自动化任务。

[![status](https://img.shields.io/badge/status-implemented-brightgreen)](#) [![license](https://img.shields.io/badge/license-MIT-blue)](LICENSE) [![focus](https://img.shields.io/badge/this--phase-Excel-green)](spec/excel-tasks.md)

---

## 这是什么

COM（Component Object Model）是 Windows 上自动化 Office 的核心机制。Office 把 `Excel.Application`
暴露为 COM 服务器，支持 COM 的语言都能驱动它的**真实引擎**去读写、编辑、计算、导出文档。

本项目聚焦两种 Excel 自动化方式的对比：

- 一套**标准任务集**（Excel 的 `E01–E12`），作为统一验收口径。
- 两种**自动化方式**（pywin32 COM / VBA）各自实现，性能对比。
- 一套**benchmark 框架**，测量不同文档规模下的操作延迟。

## 自动化方式

| 方式 | 接入方式 | 会话模型 | 备注 |
|------|----------|----------|------|
| **pywin32** | Python → 跨进程 COM IPC | 持久会话，打开一次多次操作 | 每属性一次 IPC 往返 |
| **VBA** | Python → `Application.Run` → 进程内 VBA | 持久会话，打开一次多次操作 | 零 IPC，最快 |

## 本期范围

- ✅ **Excel** —— 标准任务集 `E01–E12`、VBA 实现、benchmark（**本期重点**）
- 🔜 Word / PowerPoint —— 仅进 [roadmap](docs/roadmap.md)
- 🔜 AI agent 封装层 —— 仅进 roadmap
- ⚠️ **运行时目标**：Windows + **sandbox 内 Wine 加载 Office**，见 [wine-sandbox-runtime](docs/wine-sandbox-runtime.md)

## 目录导航

```
docs/      概念与框架文档（COM 概述、Excel 对象模型、PK 框架、基准规范、Wine 运行时、路线图）
spec/      标准任务集 E01–E12 + 功能覆盖矩阵
languages/ VBA 实现
skills/    Excel editor skill（pywin32 + VBA 后端）
benchmarks/ 基准跑分结果
assets/    共享样例数据
```

核心文档：
- [docs/00-com-overview.md](docs/00-com-overview.md) — COM 与对象模型基础
- [docs/01-excel-object-model.md](docs/01-excel-object-model.md) — Excel 对象树详解
- [spec/excel-tasks.md](spec/excel-tasks.md) — **标准任务集（主心骨）**
- [docs/pk-framework.md](docs/pk-framework.md) — 四维 PK 评分框架
- [docs/wine-sandbox-runtime.md](docs/wine-sandbox-runtime.md) — Wine+sandbox 运行时

## License

[MIT](LICENSE)
