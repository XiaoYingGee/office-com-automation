# Office COM Automation —— 多语言操作 Office 的"比武场"

> 用 **COM** 调用 Microsoft Office（Excel / Word / PowerPoint）的原生编辑接口，
> 用 **3 种语言**各实现同一组标准任务，再做**四维能力 PK**。

[![status](https://img.shields.io/badge/status-implemented-brightgreen)](#) [![license](https://img.shields.io/badge/license-MIT-blue)](LICENSE) [![focus](https://img.shields.io/badge/this--phase-Excel-green)](spec/excel-tasks.md)

---

## 这是什么

COM（Component Object Model）是 Windows 上自动化 Office 的核心机制。Office 把 `Excel.Application`、
`Word.Application`、`PowerPoint.Application` 暴露为 COM 服务器，任何支持 COM 的语言都能驱动它的**真实引擎**
去读写、编辑、计算、导出文档。

本项目把"同一件事用不同语言做一遍"组织成一个可对比的实验场：

- 一套**语言无关的标准任务集**（Excel 的 `E01–E12`），作为所有语言的统一验收口径。
- 三个**参赛语言**各自按任务集实现，互不耦合。
- 一套**四维 PK 框架**，横向比较谁更强、谁更省、谁更稳。

## 参赛语言（3）

| 语言 | COM 接入方式 | 绑定 | 备注 |
|------|--------------|------|------|
| **VBA** | 宿主内对象模型 | 早绑定 | Office 内嵌，"母语" |
| **C++** | `#import` + smart pointers | 早绑定 | 最底层、最可控 |
| **Rust** | `IDispatch` via `windows` crate | 晚绑定 | 零运行时、内存安全、RAII |

## 本期范围

- ✅ **Excel** —— 标准任务集 `E01–E12`、文档、三语言实现（**本期重点**）
- 🔜 Word / PowerPoint —— 仅进 [roadmap](docs/roadmap.md)
- 🔜 AI agent 封装层 —— 仅进 roadmap
- ⚠️ **运行时目标**：Windows + **sandbox 内 Wine 加载 Office**，见 [wine-sandbox-runtime](docs/wine-sandbox-runtime.md)

## 目录导航

```
docs/      概念与框架文档（COM 概述、Excel 对象模型、PK 框架、基准规范、Wine 运行时、路线图）
spec/      标准任务集 E01–E12 + 功能覆盖矩阵
languages/ 三语言各自的实现（vba/cpp/rust）
benchmarks/ 基准跑分方法与结果落盘
assets/    共享样例数据
```

核心文档：
- [docs/00-com-overview.md](docs/00-com-overview.md) — COM 与对象模型基础
- [docs/01-excel-object-model.md](docs/01-excel-object-model.md) — Excel 对象树详解
- [spec/excel-tasks.md](spec/excel-tasks.md) — **标准任务集（主心骨）**
- [docs/pk-framework.md](docs/pk-framework.md) — 四维 PK 评分框架
- [docs/wine-sandbox-runtime.md](docs/wine-sandbox-runtime.md) — Wine+sandbox 运行时

## PK 总览（占位 · 待跑分）

四个维度：**功能覆盖度 · 性能基准 · 代码简洁度/可维护性 · 环境/部署成本（含 Wine 可用性）**。

| 语言 | 功能覆盖 | 性能(E03) | 简洁度 | 部署成本 | Wine 可用性 |
|------|:---:|:---:|:---:|:---:|:---:|
| VBA | 12/12 | — | — | — | — |
| C++ | 12/12 | — | — | — | — |
| Rust | — | — | — | — | — |

> 评分规则见 [docs/pk-framework.md](docs/pk-framework.md)；矩阵明细见 [spec/capability-matrix.md](spec/capability-matrix.md)。

## License

[MIT](LICENSE)
