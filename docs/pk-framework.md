# PK 框架 · 四维能力评分

> 本项目的核心产出之一：用统一口径对六种语言做横向能力对比。本文定义**评什么、怎么评、怎么打分**。
> 实际跑分在各语言实现完成后进行；本期只固化框架与模板。

## 评分对象

所有语言实现**同一套** Excel 标准任务集 [`E01–E12`](../spec/excel-tasks.md)。PK 不比"谁能多做花活"，
而是比"做同一组规定动作时谁更强/更省/更稳"。

## 四个维度

### D1 · 功能覆盖度（Functional Coverage）

能把 `E01–E12` 实现到什么程度。逐任务三态评级，汇总成覆盖率。

| 评级 | 含义 | 计分 |
|---|---|---|
| ✅ 完整 | 任务全部达标 | 1.0 |
| ⚠️ 部分 | 能做但有缺陷/绕路/不稳定 | 0.5 |
| ❌ 无法 | 该语言/运行时下做不到 | 0.0 |

`D1 得分 = Σ任务评级 / 12`。明细落在 [capability-matrix.md](../spec/capability-matrix.md)。

### D2 · 性能基准（Performance）

统一基准任务的 wall-clock 耗时，**越低越好**。主项 `E03`（批量区域写入），另测 `E03-naive`
（逐格循环）以暴露封送开销差异。测量规范见 [benchmark-spec.md](benchmark-spec.md)。

- 记录：中位数耗时（重复 N 次）、机器规格、Office 版本、运行时（native / Wine）。
- 评分：对每项基准做**相对归一化**——最快者 = 1.0，其余 = `最快耗时 / 本语言耗时`。

### D3 · 代码简洁度 / 可维护性（Maintainability）

实现同一任务的代价与可读性。

| 子项 | 测量 |
|---|---|
| 代码量 | 完成 `E01–E12` 的有效 LOC（去注释/空行） |
| 可读性 | 主观 1–5 分（命名、结构、样板多少） |
| 错误处理 | 是否覆盖 COM 异常、资源释放是否健壮（与 `E12` 联动） |

评分：LOC 做相对归一化（越少越好）+ 可读性/错误处理打分，按权重合成。

### D4 · 环境 / 部署成本（Deployability）

落地一个能跑的环境有多贵——**本项目的关键维度**，因为目标 runtime 是 sandbox 内 Wine。

| 子项 | 测量 |
|---|---|
| 依赖项 | 需要哪些运行时/包/SDK（如 .NET、pywin32、WSH） |
| 安装步骤数 | 从空环境到能跑的步骤复杂度 |
| 是否需装 Office | 全员需要（COM 前提），记是否还需额外组件 |
| CI 可行性 | 能否在无人值守/headless 跑 |
| **Wine 可用性** ★ | 在 sandbox+Wine 下能否跑通：native ✅ / Wine ✅ / Wine ⚠️ / Wine ❌ |

> **Wine 可用性是加权重点**：某语言原生再强，若其 COM 栈在 Wine 下不可用，则在目标 runtime 上价值大打折扣。
> 详见 [wine-sandbox-runtime.md](wine-sandbox-runtime.md)。

## 总分合成

```
Total = w1·D1 + w2·D2 + w3·D3 + w4·D4          (Σwi = 1)
```

默认权重（可在跑分前调整并记录）：

| 维度 | 权重 | 理由 |
|---|---|---|
| D1 功能覆盖度 | 0.35 | 能做到才有意义 |
| D2 性能基准 | 0.20 | 批处理场景重要但非唯一 |
| D3 简洁度/可维护性 | 0.15 | 影响长期成本 |
| D4 部署成本（含 Wine） | 0.30 | 目标 runtime 受 Wine 强约束 |

## 打分模板（每语言一份，跑分阶段填）

```markdown
## <语言> PK 记分卡
- 环境：Office <版本> / runtime <native|wine 版本> / 机器 <CPU·RAM>
- D1 功能覆盖：__/12 = __     （明细见 capability-matrix）
- D2 性能：E03 中位数 __ ms（N=__）；E03-naive __ ms；归一化 __
- D3 简洁度：总 LOC __；可读性 _/5；错误处理 _/5
- D4 部署：依赖[…]；步骤数 __；Wine 可用性 [✅/⚠️/❌]；CI [可/否]
- 加权总分：__
- 备注/坑：…
```

## 公平性约定

- 同一台机器、同一 Office 版本、同一份样例数据（见 `assets/sample-data`）。
- 每语言都允许使用其**惯用最优写法**（如批量封送、关闭 ScreenUpdating）——比"地道实现"，不比"故意写差"。
- 所有结果连同环境信息一并落盘到 [`benchmarks/results/`](../benchmarks/)，可复现。

---

← [benchmark-spec.md](benchmark-spec.md) ｜ [capability-matrix.md](../spec/capability-matrix.md)
