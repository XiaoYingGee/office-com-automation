# Wine + Sandbox 运行时

> **关键约束**：本项目最终 runtime 是 **Windows 语义环境**，并计划在 **sandbox 内用 Wine 加载 Office 组件**运行。
> Wine 下不同语言的 COM 栈支持度差异巨大，直接决定 PK 的 D4（部署成本）维度。本文记录目标运行时、装配思路、
> 逐语言可用性预判与验证清单。
>
> ⚠️ 下表为**待验证的工程预判**，非实测结论。实现阶段需用 `E01–E12` 在真实 Wine prefix 上逐项验证并回填。

## 1. 为什么是 Wine

目标是让这套 Office 编辑能力跑在**受控 sandbox**里（后续供 AI agent 调用）。在 Linux/容器化 sandbox 中
跑真实 Office COM，主流路径是 **Wine**：用 Wine prefix 装 Windows 版 Office，再在其中跑各语言的 COM 客户端。

关键事实：**COM 进程内/本地服务器机制（`CoCreateInstance`、`IDispatch`、ProgID→CLSID 注册表查找）在
Wine 中有实现**，因此"Wine 里 New 一个 Excel.Application"在原理上可行；难点在于**各语言运行时本身能否在
Wine 跑、以及 Office 在 Wine 的稳定性**。

## 2. 运行时装配（实现阶段细化）

高层步骤（占位，待落地为可复现脚本）：

1. 准备 Wine prefix（64 位，较新稳定版 Wine / 或 Proton 衍生）。
2. 在 prefix 内安装 Windows 版 Microsoft Office（Excel 必需）。
3. 首启 Excel 完成注册，确认 `Excel.Application` ProgID 在 prefix 注册表可解析。
4. 安装目标语言运行时（见各语言行）。
5. 跑 `E01`（创建/保存 .xlsx）作为 smoke test；通过后再跑全套。

> 本期不产出安装脚本；仅约定流程。实现阶段把可工作的配置固化为 `languages/<lang>/` 下的 setup 说明 + Dockerfile/脚本。

## 3. 逐语言 Wine 可用性预判（待验证）

| 语言 | COM 客户端运行时在 Wine | 预判 | 主要风险点 |
|---|---|---|---|
| **VBA** | 跑在 Office 进程内（不需外部运行时） | 🟢 较乐观 | 依赖 Office 在 Wine 的稳定性；宏安全设置 |
| **JScript (WSH)** | `wscript`/`cscript` 由 Wine 内置提供 | 🟡 中 | Wine 的 WSH 实现完整度；`CreateObject` 行为 |
| **PowerShell** | 需 Windows PowerShell 或 PS7 在 Wine 运行 | 🟡 中 | PS 在 Wine 历来不稳；PS7 自带运行时更可控 |
| **C++** | 编译产物为原生 PE，调用 `ole32`/`oleaut32` | 🟡 中 | Wine 的 OLE/automation 覆盖；构建链 |
| **C# / .NET** | 需 .NET 运行时在 Wine（.NET Framework via Mono/wine-mono 或 .NET Core PE） | 🟠 偏谨慎 | Interop + 运行时双重依赖；wine-mono 与真 .NET 差异 |
| **Python (pywin32)** | 需 Windows 版 Python + pywin32 在 Wine | 🟠 偏谨慎 | pywin32 底层 PE 扩展在 Wine 的兼容性 |

图例：🟢相对乐观 🟡不确定/需实测 🟠风险较高。**以上均需 `E01` smoke test 实证后回填到
[capability-matrix.md](../spec/capability-matrix.md) 的"Wine 可用性"列。**

## 4. Wine 与原生 Windows 的常见差异点（排查清单）

实现/验证时重点盯这些容易在 Wine 上出问题的地方：

- **注册表 / ProgID 注册**：Office 是否在 prefix 内正确注册了自动化 CLSID。
- **DCOM / 进程外激活**：本地 out-of-proc 服务器激活路径在 Wine 的行为。
- **字体**：缺字体会影响渲染、列宽 AutoFit、PDF 导出结果。
- **PDF 导出（E10）**：`ExportAsFixedFormat` 依赖的组件在 Wine 是否可用。
- **VBA 宏执行（E11）**：宏安全策略与脚本引擎在 Wine 的可用性。
- **资源释放（E12）**：Wine 下 `EXCEL.EXE` 进程退出/僵尸行为是否与原生一致。
- **文件路径**：Windows 路径 ↔ Wine 映射（`Z:\`、drive_c）。

## 5. 验证清单（每语言在 Wine 上跑）

```text
[ ] smoke: E01 创建并保存 .xlsx，能在 host 用 openpyxl/解压校验
[ ] E02 单格读写正确
[ ] E03 批量写入正确且不崩
[ ] E04 公式重算结果正确
[ ] E05 格式（颜色/数字格式）落盘正确
[ ] E10 导出 PDF 成功且非空
[ ] E11 运行 VBA 宏（若该语言走这条路）
[ ] E12 收尾后无残留 EXCEL.EXE 进程
[ ] 记录 Wine 版本 / Office 版本 / 失败项与 workaround
```

## 6. 回填约定

每语言验证后，更新两处：
1. [spec/capability-matrix.md](../spec/capability-matrix.md) 的 **Wine 可用性**列与各任务 Wine 态。
2. 本文第 3 节表格的"预判"改为"实测"，附 Wine 版本与关键坑。

---

← [benchmark-spec.md](benchmark-spec.md) ｜ [roadmap.md](roadmap.md)
