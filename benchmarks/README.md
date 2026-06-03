# Benchmarks · 基准跑分

> 服务 PK 框架的 **D2 性能维度**。测量规范见 [../docs/benchmark-spec.md](../docs/benchmark-spec.md)。
> 本期只铺方法与目录，**不含实际跑分数据**（P3 阶段填）。

## 怎么跑

1. 确认环境：同一台机器、同一 Office 版本、记录 native / Wine。
2. 每语言的 `run` 入口都支持"基准模式"，跑 `B1 / B1-naive / B2 / B3`，输出每配置的中位数耗时。
3. 数据来自 [`../assets/sample-data`](../assets/sample-data)，用固定 seed 复现。
4. 把结果写入 `results/<language>.csv`（格式见下）。

## 基准任务（摘自 benchmark-spec）

| 编号 | 对应任务 | 说明 |
|---|---|---|
| B1 | E03 | 批量区域写入（地道，一次性二维数组封送）—— 主项 |
| B1-naive | E03 | 逐格循环写入 —— 暴露封送开销 |
| B2 | E04 | 写公式 + 强制重算 |
| B3 | E10 | 导出 PDF |

规模档：S(1k×10) / M(10k×10) / L(50k×20)。

## 结果格式

`results/<language>.csv`：

```csv
task,scale,binding,runtime,median_ms,runs,notes
B1,M,early,native-windows,,5,
B1-naive,M,early,native-windows,,5,
```

附一份 `results/<language>.env.yaml` 记录环境元数据（office_version / runtime / cpu / ram / seed / switches）。

## 汇总产物（P3）

- 分组柱状图：x=语言，分面=规模档。
- `B1` vs `B1-naive` 倍率对比（封送开销可视化）。
- 同语言 native vs Wine 对照（Wine 性能代价）。

## 目录

```
benchmarks/
├── README.md          # 本文件
└── results/           # 跑分落盘（本期空，.gitkeep 占位）
```
