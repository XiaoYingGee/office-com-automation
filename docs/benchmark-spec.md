# 基准规范 · 性能怎么测才算数

> 服务于 PK 框架的 **D2 性能维度**。统一测量方法，保证六语言结果可比、可复现。
> 本期定义规范；实际跑分在实现阶段。

## 1. 基准任务

| 编号 | 任务 | 目的 |
|---|---|---|
| **B1 = E03** | 批量区域写入 `N×M` 单元格（一次性二维数组封送） | 主项：地道写法的吞吐 |
| **B1-naive** | 同样 `N×M`，但**逐格循环**写 `Cells(r,c)` | 对照：暴露跨进程封送开销 |
| B2 = E04 | 写 `N` 条公式 + 强制重算 | 计算引擎驱动开销 |
| B3 = E10 | 导出含 `N×M` 数据的工作簿为 PDF | 引擎输出开销 |

> `B1` vs `B1-naive` 的倍率，是 [00-com-overview.md](00-com-overview.md) "批量优于循环"原则的量化证据，
> 也是各语言最具区分度的指标。

## 2. 数据规模

固定规模档位，三档都跑，观察伸缩性：

| 档 | 行 × 列 | 单元格数 |
|---|---|---|
| S | 1,000 × 10 | 1万 |
| M | 10,000 × 10 | 10万 |
| L | 50,000 × 20 | 100万 |

数据内容：固定伪随机但**可复现**（同一 seed 生成），见 `assets/sample-data`。`B1-naive` 在 L 档可能极慢，
允许只跑到 M 档并注明。

## 3. 测量方法

- **计时口径**：只计"开始写入 → 写入完成（含必要的 flush/重算）"的 wall-clock，**不含** Excel 启动和文件保存
  （保存单独记，避免磁盘噪声混入）。
- **重复**：每配置跑 `N=5` 次，**丢弃首次**（冷启动/JIT/早绑定缓存预热），取剩余**中位数**。
- **预热**：早绑定语言（C#/Python-EnsureDispatch/C++）先做一次小写入预热类型缓存。
- **隔离**：每次用全新 Workbook；进程间确保上一轮 `EXCEL.EXE` 已退出（与 `E12` 联动）。
- **统一开关**：测主项时按地道写法设 `ScreenUpdating=False`、`Calculation=Manual`、`DisplayAlerts=False`；
  并在结果中**记录所用开关**（公平前提是写法公开）。

## 4. 必须记录的环境元数据

每条结果都要附带，否则不可比：

```yaml
language:        # powershell | python | csharp | vba | jscript | cpp
binding:         # early | late
office_version:  # 例 Office 2021 / Microsoft 365 16.x
runtime:         # native-windows | wine-<version>
os:              # Windows 11 26xxx / sandbox 描述
cpu:             # 型号
ram_gb:
seed:            # 数据生成种子
switches:        # ScreenUpdating/Calculation/... 实际取值
```

## 5. 结果落盘格式

存到 [`benchmarks/results/`](../benchmarks/)，每语言一个文件，建议 CSV/JSON：

```csv
task,scale,binding,runtime,median_ms,runs,notes
B1,M,early,native-windows,,5,
B1-naive,M,early,native-windows,,5,
B2,M,early,native-windows,,5,
B3,M,early,native-windows,,5,
```

> 空 `median_ms` = 待跑。本期只铺格式，不填数。

## 6. 归一化与呈现

- 每个 `(task, scale)` 组内，最快语言 = `1.00`，其余 = `min / self`。
- 汇总图：分组柱状图（x=语言，分面=规模档），以及 `B1` vs `B1-naive` 倍率对比。
- Wine vs native 同语言对照单列一张，量化 Wine 的性能代价。

---

← [pk-framework.md](pk-framework.md) ｜ [wine-sandbox-runtime.md](wine-sandbox-runtime.md)
