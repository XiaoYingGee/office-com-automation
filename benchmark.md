# Benchmark — Excel 自动化 4 后端对比

> Date: 2026-06-09
> 统一 11 op，9 项测试，4 后端 (Rust COM / OpenXML .NET / pywin32 / VBA)
> 环境: Windows 11 + Excel (Office 16.0), Python 3.10, 5 measured rounds + 1 warmup

---

## 测试项 (9 项)

| # | 测试 | 操作 | 考察点 |
|---|------|------|--------|
| B1 | Cell Write | 写 5 个单元格 | 基础写入延迟 |
| B2 | Bulk Write | 写 100×10 区域 | 批量写入 |
| B3 | Read Cell | 读取单元格 | 读取延迟 |
| B4 | Clear Range | 清除 1000 单元格 | 批量清除 |
| B5 | Inspect Workbook | 读工作簿结构 | 读密集 (IPC 差异最大) |
| B6 | Batch 10 Writes | 连续 10 次写 | 多次操作累积 |
| B7 | Format 5 Cells | bold+size+color | 多属性设置 |
| B8 | Insert 5 Rows | 插入 5 行 | 结构操作 |
| B9 | Sheet Ops | Add+Rename+Delete | 工作表管理 |

---

## 测试结果

延迟 = total ms per test，**越低越快**，每项最快用粗体。

| # | 测试 | pywin32 | vba | openxml | rust |
|---|------|------:|------:|------:|------:|
| B1 | Cell Write (5 cells) | 509 | **408** | 1647 | _pending_ |
| B2 | Bulk Write (100×10) | 37 | **28** | 352 | _pending_ |
| B3 | Read Cell | 36 | **5** | 317 | _pending_ |
| B4 | Clear Range (1000) | **20** | 21 | 336 | _pending_ |
| B5 | Inspect Workbook | 27 | **24** | 318 | _pending_ |
| B6 | Batch 10 Writes | 135 | **9** | 3181 | _pending_ |
| B7 | Format 5 Cells | 183 | **36** | 1621 | _pending_ |
| B8 | Insert 5 Rows | 38 | **20** | 311 | _pending_ |
| B9 | Sheet Add+Rename+Delete | 51 | **15** | 918 | _pending_ |

> Rust 后台单独采集中 (每 op spawn Excel 进程，按设计很慢)，回来后补入本表。

**实测排名 (3/4): 🥇 VBA → 🥈 pywin32 → 🥉 OpenXML**

---

## 架构

```
pywin32:  Python ──COM IPC (每属性1次)──→ Excel 进程
vba:      Python ──App.Run (1次)──→ Excel 进程内 VBA 执行
openxml:  [启动 exe] ──读+写整个 .xlsx──→ 文件 (无 Excel)   ← 每 op 固定开销
rust:     [启动 exe] ──COM──→ Excel 进程 ──关闭── (每 op)
```

---

## 关键发现

### 1. VBA 在 IPC 密集型操作上碾压 (符合预期)

VBA 赢 8/9。差距最大处正是跨进程 COM IPC 主导的场景:
B6 Batch (135→9, **15×**)、B7 Format (183→36, **5×**)、B3 Read (36→5, **7×**)。
印证: VBA 进程内执行零 per-property IPC，pywin32 每属性一次 COM 往返。

### 2. OpenXML 实测最慢，与原预期相反 ⚠️

原预期 OpenXML 为 🥉/"快速文件 I/O"，**实测在所有 9 项中均最慢**，差距巨大 (B6 = 3181ms)。

根因: OpenXML 后端是 **standalone exe 每 op 调用一次**，每次都要
*启动进程 + 读整个 .xlsx + 改 + 整文件写回*。~300ms 固定开销 (见 B3/B4/B5/B8 均 ≈310–340ms)
压倒"不启动 Excel"的收益。B6 (10 次写) 最差，因为重复整文件读写 10 次。

OpenXML 的优势仅在**大批量、一次性、无需实时 Excel** 的变换场景才体现，
本 benchmark 的小粒度高频 op 反而是其最差场景。

---

## 复现

```powershell
cd skills\excel-editor\scripts
python benchmark.py --all `
  --rust-exe "D:/Workspace/office-com-automation/languages/rust/target/release/excel-ops.exe" `
  --openxml-exe "D:/Workspace/office-com-automation/languages/openxml/bin/Release/net10.0/ExcelOps.exe" `
  --rounds 5
```

> 注意: exe 路径用正斜杠绝对路径，避免 shell 反斜杠转义导致 "exe not found"。
> VBA 后端需 Excel 信任中心开启 "Trust access to the VBA project object model"
> (注册表 `HKCU\Software\Microsoft\Office\16.0\Excel\Security\AccessVBOM = 1`)。
