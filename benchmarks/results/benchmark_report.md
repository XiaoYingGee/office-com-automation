# Excel MCP Benchmark Report

> Date: 2026-06-09
> Machine: Windows 11 Pro
> Timing: `elapsed_ms` from server-side Stopwatch/perf_counter (pure COM execution, no LLM overhead)

## excel-pywin32

| # | Prompt | 文件 | open | execute | save | **总计** | 结果 |
|---|--------|------|------|---------|------|---------|------|
| P1 | 基础读写 | empty | 88 ms | **14 ms** | 28 ms | **130 ms** | ✅ |
| P2 | 批量写入 1000行 | empty | 13 ms | **76 ms** | 24 ms | **113 ms** | ✅ |
| P3 | 条件格式 25k行 | medium | 3,266 ms | **24,188 ms** | 2,939 ms | **30,393 ms** | ✅ |
| P4 | 跨表汇总 | medium | — | **149 ms** | 2,874 ms | **3,023 ms** | ✅ |
| P5 | 图表创建 | medium | — | **2,440 ms** | 2,834 ms | **5,274 ms** | ✅ |
| P6 | 查找替换+公式 | medium | — | **1,691 ms** | 2,889 ms | **4,580 ms** | ✅ |
| P7 | 大文件操作 | large | 6,279 ms | **24 ms** | 5,906 ms | **12,209 ms** | ✅ |
| P8 | 错误恢复 | empty | 45 ms | **17 ms** | 27 ms | **89 ms** | ✅ |
| P9 | 复杂格式化 | empty | 11 ms | **252 ms** | 27 ms | **290 ms** | ✅ |
| P10 | 数据透视表 | medium | 57 ms | **737 ms** | 3,094 ms | **3,888 ms** | ✅ |

**Execute 总计: 29,588 ms | 含I/O总计: 59,989 ms (≈60s)**

## excel-csharp

| # | Prompt | 文件 | open | execute | save | **总计** | 结果 |
|---|--------|------|------|---------|------|---------|------|
| P1 | 基础读写 | empty | — | **170 ms** | 57 ms | **227 ms** | ✅ |
| P2 | 批量写入 1000行 | empty | — | **49 ms** | 22 ms | **71 ms** | ✅ |
| P3 | 条件格式 25k行 | medium | 3,718 ms | **27,010 ms** | 3,033 ms | **33,761 ms** | ✅ |
| P4 | 跨表汇总 | medium | — | **167 ms** | 2,970 ms | **3,137 ms** | ✅ |
| P5 | 图表创建 | medium | — | **2,571 ms** | 2,968 ms | **5,539 ms** | ✅ |
| P6 | 查找替换+公式 | medium | — | **1,424 ms** | 2,940 ms | **4,364 ms** | ✅ |
| P7 | 大文件操作 | large | 6,691 ms | **66 ms** | 6,299 ms | **13,056 ms** | ✅ |
| P8 | 错误恢复 | empty | — | **39 ms** | 63 ms | **102 ms** | ✅ |
| P9 | 复杂格式化 | empty | — | **530 ms** | 59 ms | **589 ms** | ✅ |
| P10 | 数据透视表 | medium | — | **831 ms** | 3,232 ms | **4,063 ms** | ✅ |

**Execute 总计: 32,857 ms | 含I/O总计: 64,909 ms (≈65s)**

## 对比总结

| Prompt | pywin32 exec | csharp exec | 倍率 | 赢家 | 分析 |
|--------|:---:|:---:|:---:|:---:|------|
| P1 基础读写 | 14 ms | 170 ms | 12× | **pywin32** | C# 逐cell写 vs Python批量Range |
| P2 批量写入 | 76 ms | 49 ms | 0.6× | **csharp** | C# object[,]数组直写SAFEARRAY |
| P3 条件格式 | 24,188 ms | 27,010 ms | 1.1× | **pywin32** | 两者都被逐行COM IPC卡死 |
| P4 跨表汇总 | 149 ms | 167 ms | 1.1× | **pywin32** | 基本持平(数据读取+内存聚合) |
| P5 图表创建 | 2,440 ms | 2,571 ms | 1.1× | **pywin32** | COM图表对象创建开销接近 |
| P6 查找替换 | 1,691 ms | 1,424 ms | 0.8× | **csharp** | Replace+逐行公式，C#略快 |
| P7 大文件追加 | 24 ms | 66 ms | 2.8× | **pywin32** | 简单操作pywin32开销更低 |
| P8 错误恢复 | 17 ms | 39 ms | 2.3× | **pywin32** | Python try/except更轻量 |
| P9 复杂格式化 | 252 ms | 530 ms | 2.1× | **pywin32** | 逐cell循环C#动态绑定开销大 |
| P10 透视表 | 737 ms | 831 ms | 1.1× | **pywin32** | 持平 |

### 性能档位分布

```
            pywin32                         csharp
< 100ms:    P1(14) P2(76) P7(24)           P2(49) P7(66) P8(39)
            P8(17)
100-999ms:  P4(149) P9(252) P10(737)       P1(170) P4(167) P9(530) P10(831)
1-5s:       P5(2440) P6(1691)              P5(2571) P6(1424)
> 20s:      P3(24188)                      P3(27010)
```

### 关键发现

1. **真实 COM 执行时间远小于之前记录的"端到端"时间** — P1 实际仅需 14ms，之前记录 91s 是 LLM 推理 + 网络往返
2. **瓶颈在逐行 COM 调用** — P3 对 25k 行逐行设置 `Interior.Color`，两个后端都需 ~25s，占总执行时间 80%+
3. **批量写入 C# 更快** — P2 中 `object[,]` 2D 数组写入 49ms vs Python 的 76ms (SAFEARRAY 直通)
4. **简单操作 pywin32 更快** — Python 的 COM dispatch 对简单属性访问开销更低（无动态编译开销）
5. **文件 I/O 时间固定** — medium 文件 save ≈3s，large 文件 open/save ≈6s，与后端无关（Excel 本身开销）
6. **C# 的 CSharpCodeProvider 编译开销** — 未含在 execute 计时中（~200ms/次），但若含编译失败重试则额外 +200-400ms

### 结论

| 维度 | pywin32 | csharp | 说明 |
|------|:---:|:---:|------|
| Execute 总时间 | **29.6s** | 32.9s | pywin32 快 10% |
| 含I/O总时间 | **60.0s** | 64.9s | I/O开销相同，差距来自 execute |
| 首次成功率 | **10/10** | 10/10 | 本轮无编译错误（已熟悉API） |
| 代码简洁度 | ⭐⭐⭐⭐⭐ | ⭐⭐⭐ | Python 无需 dynamic cast |
| 批量写入 | 76ms | **49ms** | C# 2D array 优势 |
| 单cell操作 | **14ms** | 170ms | pywin32 dispatch更轻 |
