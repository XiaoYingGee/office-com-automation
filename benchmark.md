# Excel MCP Server Benchmark

> 测量日期：2026-06-09
> 环境：Windows 11 Pro + Microsoft Excel (Office 16.0)
> 计时方法：server-side `time.perf_counter()` (Python) / `Stopwatch` (C#)，纯 COM 执行时间

---

## 1. 概述

本 benchmark 对比两个 MCP Server 后端驱动 Excel 的真实 COM 执行性能：

| 后端 | 机制 | CodeAct |
|------|------|---------|
| **excel-pywin32** | Python → 跨进程 COM IPC | `exec()` 执行 Python 代码 |
| **excel-csharp** | C# .NET 4.8 → 跨进程 COM | `CSharpCodeProvider` 动态编译执行 |

---

## 2. 测试矩阵 (10 个 Prompt)

| # | 操作 | 文件 | 大小 | 考察点 |
|---|------|------|------|--------|
| P1 | 5×5 写入 + 加粗 | bench_empty.xlsx | 5 KB | 基础读写 |
| P2 | 1000行×10列批量写入 | bench_empty.xlsx | 5 KB | 批量数据 + 格式化 |
| P3 | 50k行条件查询 + 逐行标黄 | bench_medium.xlsx | 43 MB | 大范围遍历 |
| P4 | 跨表聚合统计 | bench_medium.xlsx | 43 MB | 数据读取 + 内存计算 |
| P5 | 柱状图创建 | bench_medium.xlsx | 43 MB | Chart 对象 |
| P6 | 查找替换 + 公式写入 | bench_medium.xlsx | 43 MB | Replace + Formula |
| P7 | 大文件追加汇总行 | bench_large.xlsx | 86 MB | 大文件 I/O |
| P8 | 错误处理 (含预期失败) | bench_empty.xlsx | 5 KB | 异常恢复 |
| P9 | 月报模板 (合并/公式/条件格式) | bench_empty.xlsx | 5 KB | 复杂格式化 |
| P10 | 数据透视表 | bench_medium.xlsx | 43 MB | PivotTable |

---

## 3. 结果

### 3.1 Execute 时间对比

| Prompt | pywin32 | csharp | 倍率 | 赢家 |
|--------|--------:|-------:|:---:|:---:|
| P1 基础读写 | **14 ms** | 170 ms | 12× | pywin32 |
| P2 批量写入 | 76 ms | **49 ms** | 0.6× | csharp |
| P3 条件格式 25k行 | **24,188 ms** | 27,010 ms | 1.1× | pywin32 |
| P4 跨表汇总 | **149 ms** | 167 ms | 1.1× | pywin32 |
| P5 图表创建 | **2,440 ms** | 2,571 ms | 1.1× | pywin32 |
| P6 查找替换+公式 | 1,691 ms | **1,424 ms** | 0.8× | csharp |
| P7 大文件追加 | **24 ms** | 66 ms | 2.8× | pywin32 |
| P8 错误恢复 | **17 ms** | 39 ms | 2.3× | pywin32 |
| P9 复杂格式化 | **252 ms** | 530 ms | 2.1× | pywin32 |
| P10 透视表 | **737 ms** | 831 ms | 1.1× | pywin32 |
| **合计** | **29,588 ms** | 32,857 ms | | **pywin32** |

### 3.2 文件 I/O 时间

| 操作 | pywin32 | csharp | 说明 |
|------|--------:|-------:|------|
| Open medium (43MB) | 3,266 ms | 3,718 ms | 基本持平 |
| Open large (86MB) | 6,279 ms | 6,691 ms | 基本持平 |
| Save medium | ~2,900 ms | ~3,000 ms | 基本持平 |
| Save large | 5,906 ms | 6,299 ms | 基本持平 |

### 3.3 可视化

```
P1  基础读写     pywin32 █ 14ms             csharp ████████ 170ms
P2  批量写入     pywin32 ███ 76ms           csharp ██ 49ms
P3  条件格式     pywin32 ████████████████ 24.2s    csharp ██████████████████ 27.0s
P4  跨表汇总     pywin32 █ 149ms            csharp █ 167ms
P5  图表创建     pywin32 ██████ 2.4s        csharp ██████ 2.6s
P6  查找替换     pywin32 ████ 1.7s          csharp ████ 1.4s
P7  大文件追加   pywin32 █ 24ms             csharp █ 66ms
P8  错误恢复     pywin32 █ 17ms             csharp █ 39ms
P9  复杂格式化   pywin32 █ 252ms            csharp ██ 530ms
P10 透视表       pywin32 ██ 737ms           csharp ███ 831ms
```

---

## 4. 分析

### 4.1 pywin32 整体胜出 (8:2)

pywin32 在 8/10 项中更快，核心原因：
- Python COM dispatch 对简单属性访问开销极低（无编译步骤）
- `Range.Value2 = tuple(data)` 自动映射 SAFEARRAY，写入高效

### 4.2 csharp 在批量数组写入占优

C# `object[,]` 2D 数组直接映射 COM SAFEARRAY，在 P2 (49ms vs 76ms) 和 P6 (1424ms vs 1691ms) 中更快。

### 4.3 逐行 COM 调用是唯一瓶颈

P3 对 25k 行逐行设置 `Interior.Color`，两个后端都需 ~25s。占总执行时间的 **80%**。应优先使用 `Range` 批量操作或 `FormatConditions` 代替逐行格式化。

### 4.4 文件 I/O 与后端无关

Open/Save 耗时由 Excel 本身的文件解析/序列化决定，两后端差异 <15%。

### 4.5 C# CSharpCodeProvider 的额外开销

每次 `execute_code` 都要动态编译（~200ms），且类型系统严格可能导致编译失败需重试。pywin32 的 `exec()` 无编译开销。

---

## 5. 总结

| 维度 | pywin32 | csharp |
|------|:---:|:---:|
| Execute 总时间 | **29.6s** | 32.9s |
| 含 I/O 总时间 | **60.0s** | 64.9s |
| 成功率 | 10/10 | 10/10 |
| 代码简洁度 | ⭐⭐⭐⭐⭐ | ⭐⭐⭐ |
| 批量写入性能 | 良好 | **优秀** |
| 单操作延迟 | **极低** | 中等 |
| 编译开销 | 无 | ~200ms/次 |

**推荐**：默认使用 `excel-pywin32`。在需要大批量 2D 数组写入时 `excel-csharp` 有微弱优势，但 pywin32 的开发体验（无类型错误、无编译步骤）显著更好。

---

## 6. 复现

```bash
# 阅读测试指令
cat benchmarks/RUN_BENCHMARK.md

# 详细数据
cat benchmarks/results/benchmark_report.md
```

## 7. 历史对比 (pywin32 vs VBA)

此前的 benchmark (见本文件 git 历史) 对比了 pywin32 vs VBA 后端。VBA 在 IPC 密集操作上快 20-50×（进程内执行零 IPC），但 VBA 后端已弃用——MCP CodeAct 模式（AI 直接写代码执行）取代了预写 VBA 宏的方式。
