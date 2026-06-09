# Excel 自动化 Benchmark

> 测量日期：2026-06-09
> 环境：Windows 11 + Microsoft Excel (Office 16.0)、Python 3.10
> 方法：每后端 × 每尺寸 1 warmup + 3 measured rounds，取均值。

---

## 1. 概述

本 benchmark 对比在 Windows 上驱动 Excel 的不同自动化后端，跨**三种文档规模**测量 11 项操作的延迟。

最终对比聚焦两个**常驻会话（persistent）**后端：

| 后端 | 机制 | 会话模型 |
|------|------|----------|
| **pywin32** | Python → 跨进程 COM IPC，每属性一次往返 | 打开一次，多次操作 |
| **vba** | Python → `Application.Run` → 进程内 VBA 执行，对象模型零 IPC | 打开一次，多次操作 |

三个**无状态（stateless）**后端（Rust COM / OpenXML .NET / C++ COM）已实现并验证，但因每个操作都要重开整个文件，规模化下不具竞争力，已从最终对比中移除——理由与样本见 §6。

---

## 2. 方法

**文档规模**（由 `skills/excel-editor/scripts/gen_fixture.py` 生成，seeded 随机内容）：

| 尺寸 | 结构 | 单元格数 | 文件大小 |
|------|------|---------|---------|
| empty | 1 sheet，无数据 | 0 | ~6 KB |
| medium | 3 sheet × 1000 行 × 50 列 | 15 万 | ~1.1 MB |
| large | 10 sheet × 20000 行 × 100 列 | 2000 万 | ~145 MB |

**运行规则**：每轮将对应尺寸的 fixture 拷贝到独立工作文件（拷贝不计时），后端在副本上操作；空文档维持运行时新建。每后端 1 warmup + 3 measured rounds 取均值。

**复现**：

```powershell
cd skills\excel-editor\scripts
python gen_fixture.py                 # 生成 empty/medium/large fixture
python benchmark.py --backends pywin32,vba --size empty,medium,large --rounds 3
```

> VBA 后端需 Excel 信任中心开启 "Trust access to the VBA project object model"
> （注册表 `HKCU\Software\Microsoft\Office\16.0\Excel\Security\AccessVBOM = 1`）。

---

## 3. 测试矩阵（11 项）

| # | 测试 | 操作 | 考察点 |
|---|------|------|--------|
| B0 | Open Workbook | 打开文档 | 文件打开延迟（随文件大小放大） |
| B1 | Cell Write | 写 5 个单元格 | 基础写入 |
| B2 | Bulk Write | 写 100×10 区域 | 批量写入 |
| B3 | Read Cell | 读单元格 | 读取延迟 |
| B4 | Clear Range | 清除 1000 单元格 | 批量清除 |
| B5 | Inspect Workbook | 读工作簿结构 | 读密集（IPC 差异最大） |
| B6 | Batch 10 Writes | 连续 10 次写 | 多次操作累积 |
| B7 | Format 5 Cells | bold + size + color | 多属性设置 |
| B8 | Insert 5 Rows | 插入 5 行 | 结构操作 |
| B9 | Sheet Ops | Add + Rename + Delete | 工作表管理 |
| B10 | Merge 5 Ranges | 合并 5 个区域 | 合并单元格 |

---

## 4. 结果

延迟单位 ms，越低越快，每项更快者加粗。

### 4.1 empty 文档

| 测试 | pywin32 | vba |
|------|------:|------:|
| B0 Open Workbook | N/A（无文件） | N/A（无文件） |
| B1 Cell Write (5) | 488 | **422** |
| B2 Bulk Write (100×10) | 33 | **31** |
| B3 Read Cell | 25 | **7** |
| B4 Clear Range (1000) | 27 | **20** |
| B5 Inspect | 28 | **13** |
| B6 Batch 10 Writes | 139 | **6** |
| B7 Format 5 Cells | 172 | **39** |
| B8 Insert 5 Rows | 38 | **15** |
| B9 Sheet Add+Rename+Delete | 51 | **20** |
| B10 Merge 5 Ranges | 105 | **6** |

### 4.2 medium 文档（15 万单元格）

| 测试 | pywin32 | vba |
|------|------:|------:|
| B0 Open Workbook | **751** | 889 |
| B1 Cell Write (5) | **424** | 551 |
| B2 Bulk Write (100×10) | 26 | **17** |
| B3 Read Cell | 13 | **4** |
| B4 Clear Range (1000) | 78 | **6** |
| B5 Inspect | 84 | **4** |
| B6 Batch 10 Writes | 214 | **6** |
| B7 Format 5 Cells | 123 | **26** |
| B8 Insert 5 Rows | 46 | **13** |
| B9 Sheet Add+Rename+Delete | 52 | **15** |
| B10 Merge 5 Ranges | 204 | **4** |

### 4.3 large 文档（2000 万单元格，~145 MB）

| 测试 | pywin32 | vba |
|------|------:|------:|
| B0 Open Workbook | 18532 | **18255** |
| B1 Cell Write (5) | **452** | 527 |
| B2 Bulk Write (100×10) | 23 | **21** |
| B3 Read Cell | 16 | **2** |
| B4 Clear Range (1000) | 105 | **11** |
| B5 Inspect | 504 | **121** |
| B6 Batch 10 Writes | 147 | **28** |
| B7 Format 5 Cells | 173 | **84** |
| B8 Insert 5 Rows | 150 | **117** |
| B9 Sheet Add+Rename+Delete | 65 | **21** |
| B10 Merge 5 Ranges | 233 | **12** |

---

## 5. 分析

### 5.1 VBA 在 IPC 密集型操作上占优

VBA 在 11 项里几乎全胜。差距最大的正是跨进程 COM IPC 主导的操作：

| 测试 | medium pywin32 / vba | 倍数 |
|------|:---:|:---:|
| B6 Batch 10 Writes | 214 / 6 | **36×** |
| B10 Merge 5 Ranges | 204 / 4 | **51×** |
| B5 Inspect | 84 / 4 | **21×** |

机理：VBA 通过一次 `Application.Run` 把整批操作送入 Excel 进程内执行，对象模型访问零 IPC；pywin32 每访问一个属性就付一次跨进程 COM 往返。操作越细碎、属性访问越多，差距越大。

### 5.2 pywin32 仅在两类场景略快

- **B0 Open（medium/large）**：VBA 打开时多一次性的 VBA 模块注入开销，故 pywin32 开文件略快。
- **B1 首次 Cell Write**：pywin32 的 gencache/COM 早绑定在首写时已暖；VBA 首次 `App.Run` 有一次调用建立开销。稳态下（B6 等）VBA 反超。

### 5.3 文件大小主要冲击"打开"，而非单次操作

随文档从 empty→medium→large：

- **B0 Open 急剧放大**：empty 不适用 → medium ~0.8s → **large ~18.5s**（两后端都受 145MB 文件解析支配，与后端机制无关）。
- **单次操作几乎不随规模变化**：B1/B2/B3/B6/B10 等在三尺寸下基本恒定——因为常驻后端打开一次后，后续操作只动局部区域，与总单元格数无关。
- **少数读密集操作随规模升**：B5 Inspect 在 large 上升到 pywin32 504ms / vba 121ms（要遍历更大的 used range）。

**结论**：对常驻后端，大文档的代价集中在一次性的打开；打开之后的吞吐与小文档相当。这与无状态后端形成鲜明对比（§6）。

---

## 6. 已评估但移除的无状态后端

Rust COM、OpenXML .NET、C++ COM 三个后端均已实现全部 13 个 op 并通过验证（代码保留在 `languages/`），但因**架构不具竞争力**从对比中移除：

- **Rust / C++（无状态 COM）**：每个 op 都 `启动 Excel 进程 → 操作 → 关闭`，固定开销约 1.7–2.3s/op。
- **OpenXML（无状态文件 I/O）**：每个 op 都 `读整个 .xlsx → 改 → 整文件写回`，约 300ms+ 固定开销/op。

empty 文档下的代表性样本（与 §4 常驻后端同口径，单位 ms）：

| 测试 | pywin32 | vba | cpp | openxml | rust* |
|------|------:|------:|------:|------:|------:|
| B1 Cell Write (5) | 488 | 422 | 9677 | 1531 | 13626 |
| B6 Batch 10 Writes | 139 | 6 | 18121 | 3241 | 22701 |
| B10 Merge 5 Ranges | 105 | 6 | 10146 | 1469 | 9588 |

> *rust 列在并发负载下采集，偏高；趋势仍代表无状态架构的量级。

无状态后端比常驻后端慢 **1–2 个数量级**，且在大文档下完全不可行：每个 op 都要重新打开 145MB 文件（光打开就 ~18s，见 §5.3），整套测试将耗时数小时。这正是它们在 large 层被标为 N/A、并最终从对比中移除的原因。

无状态架构的唯一适用场景是**大批量、一次性、无需实时 Excel 会话**的变换——与本 benchmark 的小粒度高频操作正好相反。
