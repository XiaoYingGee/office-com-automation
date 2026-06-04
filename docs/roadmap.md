# Roadmap · 路线图

> 本项目分阶段推进：先把"比武场"骨架与 Excel 标准动作立起来，再逐语言实现、跑分对比，最后扩展组件并封装给 AI agent。

## 阶段总览

| 阶段 | 名称 | 范围 | 状态 |
|---|---|---|---|
| **P0** | 脚手架 + 规划 | 仓库结构、文档、`E01–E12` 任务集、PK 框架、Wine 运行时说明、分语言 plan | ✅ **本期** |
| P1 | 各语言 Excel 实现 | 3 语言各自实现 `E01–E12`，落 `languages/<lang>/src/` | ✅ 已完成 |
| P2 | Wine 运行时打通 | 在 sandbox+Wine 跑通各语言，回填 [wine-sandbox-runtime](wine-sandbox-runtime.md) 与能力矩阵 | ⏳ |
| P3 | 四维 PK 跑分 | 按 [benchmark-spec](benchmark-spec.md) 跑分，填 [pk-framework](pk-framework.md) 记分卡与 README 总览表 | ⏳ |
| P4 | Word / PowerPoint 扩展 | 复用同套方法，新增 `W01…` / `P01…` 任务集 | ⏳ |
| P5 | AI agent 封装层 | 把胜出方案封装成 agent 可调用的工具接口 | ⏳ |

## P0 · 本期交付（已完成口径）

- 公开仓库 + 目录结构 + MIT License。
- 概念文档：COM 概述、Excel 对象模型。
- 标准任务集 `E01–E12`（语言无关验收口径）。
- PK 四维框架 + 基准规范 + 能力矩阵骨架。
- Wine+sandbox 运行时说明与逐语言可用性预判。
- 三语言各自的实现 plan（README）。
- **不含**任何可运行实现代码、不跑实际跑分。

## P1 · 各语言 Excel 实现

- 每个 `languages/<lang>/` 下建 `src/`，按 README 约定实现 `E01–E12`。
- 统一 CLI/入口：能单跑某任务、能跑全套、能产出基准计时。
- 共享样例数据来自 `assets/sample-data`，结果可复现。
- 每语言补充其实测坑与释放策略。

## P2 · Wine 运行时打通

- 固化可复现的 Wine prefix + Office 安装流程（脚本/Dockerfile）。
- 逐语言跑 [验证清单](wine-sandbox-runtime.md#5-验证清单每语言在-wine-上跑)。
- 回填能力矩阵的"Wine 可用性"列。

## P3 · 四维 PK 跑分

- 统一机器跑 D2 性能；汇总 D1/D3/D4。
- 填 README 顶部 **PK 总览表** 与每语言记分卡。
- 产出对比图（性能分组柱状图、Wine vs native 对照）。

## P4 · Word / PowerPoint 扩展

- Word：`Document/Range/Paragraph/Table` 编辑任务集 `W01…`。
- PowerPoint：`Presentation/Slide/Shape` 任务集 `P01…`。
- 复用 PK 框架，按组件分别评分。

## P5 · AI agent 封装层（远期设想）

> 目标：把这套 Office 编辑能力封装成 **AI agent 可调用的工具**，跑在 sandbox+Wine runtime 上。

可能方向（待 P1–P3 结论确定技术选型后细化）：

- 选 PK 综合胜出的语言/方案作为执行后端。
- 暴露为结构化工具接口（例如 MCP server，或函数调用式 API）：
  `open / write_range / set_format / add_formula / export_pdf / run_macro / close` 等原子能力。
- 在 sandbox 内以受控方式驱动 Office，处理并发、超时、僵尸进程治理、错误回报。
- 安全：限制宏执行、隔离文件系统、资源配额。

**本设想仅记录方向，不在 P0–P3 实现目标内。**

---

← [wine-sandbox-runtime.md](wine-sandbox-runtime.md) ｜ 任务集 → [../spec/excel-tasks.md](../spec/excel-tasks.md)
