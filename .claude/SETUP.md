# Excel MCP Server — Claude Code 接入指南

## 快速安装

### 方式 1：`claude mcp add` 命令（推荐）

```bash
# 安装 pywin32 后端（推荐，性能最优）
claude mcp add excel-pywin32 -- py D:\Workspace\AI\office-com-automation\mcp_server\excel_mcp.py

# 安装 addin 后端（进程内执行，需先注册 add-in）
claude mcp add excel-addin -- py D:\Workspace\AI\office-com-automation\mcp_server\excel_mcp_addin.py

# 安装 C# 后端
claude mcp add excel-csharp -- D:\Workspace\AI\office-com-automation\mcp_server_csharp\ExcelMcp\bin\Release\net48\ExcelMcp.exe
```

加 `--scope user` 注册为全局（所有项目可用），不加则为当前项目级：

```bash
# 全局安装
claude mcp add excel-pywin32 --scope user -- py D:\Workspace\AI\office-com-automation\mcp_server\excel_mcp.py
```

安装后重启 Claude Code，输入 `/mcp` 验证 server 状态。

### 方式 2：手动编辑 settings.json

如果 `claude mcp add` 不可用，可手动编辑配置文件：

- 项目级：`<project>/.claude/settings.json`
- 全局：`~/.claude/settings.json`

```json
{
  "mcpServers": {
    "excel-pywin32": {
      "command": "py",
      "args": ["D:\\Workspace\\AI\\office-com-automation\\mcp_server\\excel_mcp.py"]
    }
  }
}
```

### 方式 3：本 repo 内直接使用

Clone 本项目后在项目目录打开 Claude Code，MCP servers 已配好自动启动：

```bash
cd D:\Workspace\AI\office-com-automation
claude
```

## 前提条件

- Windows 10/11
- Microsoft Excel 已安装（Office 2016+）
- Python 3.10+ + `pip install pywin32 mcp`
- （可选）C# server 需要 .NET Framework 4.8 运行时

## 多后端同时启用

```bash
claude mcp add excel-pywin32 -- py D:\Workspace\AI\office-com-automation\mcp_server\excel_mcp.py
claude mcp add excel-csharp -- D:\Workspace\AI\office-com-automation\mcp_server_csharp\ExcelMcp\bin\Release\net48\ExcelMcp.exe
claude mcp add excel-dotnet -- dotnet run --project D:\Workspace\AI\office-com-automation\mcp_server_dotnet\ExcelMcpDotnet.csproj
```

### 后端对比

| 后端 | 语言 | 需要 | 协议 | 推荐场景 |
|------|------|------|------|----------|
| `excel-pywin32` | Python | `pip install pywin32 mcp` | FastMCP SDK | **默认推荐** |
| `excel-csharp` | C# | 无额外安装 | 手写 JSON-RPC | 无 Python 环境时 |
| `excel-dotnet` | C# | .NET 8 SDK | 官方 MCP SDK | 正式/生产环境 |
| `excel-addin` | Python | add-in 注册 | FastMCP SDK | 进程内零 IPC |

或手动 settings.json：

```json
{
  "mcpServers": {
    "excel-pywin32": {
      "command": "py",
      "args": ["D:\\Workspace\\AI\\office-com-automation\\mcp_server\\excel_mcp.py"]
    },
    "excel-csharp": {
      "command": "D:\\Workspace\\AI\\office-com-automation\\mcp_server_csharp\\ExcelMcp\\bin\\Release\\net48\\ExcelMcp.exe",
      "args": []
    },
    "excel-dotnet": {
      "command": "dotnet",
      "args": ["run", "--project", "D:\\Workspace\\AI\\office-com-automation\\mcp_server_dotnet\\ExcelMcpDotnet.csproj"]
    }
  }
}
```

## 卸载

```bash
claude mcp remove excel-pywin32
claude mcp remove excel-addin
claude mcp remove excel-csharp
claude mcp remove excel-dotnet
```

## 权限配置（免确认）

在 `settings.json` 或 `settings.local.json` 中添加：

```json
{
  "permissions": {
    "allow": [
      "mcp__excel-pywin32__excel_open",
      "mcp__excel-pywin32__excel_save",
      "mcp__excel-pywin32__excel_inspect",
      "mcp__excel-pywin32__excel_inspect_sheet",
      "mcp__excel-pywin32__excel_execute_code",
      "mcp__excel-pywin32__excel_execute_actions"
    ]
  }
}
```

## 使用方式

接入后 Claude Code 可直接操作 Excel：

```
> 打开 report.xlsx，在 A1 写入 Hello

> 读取 Sheet1 的 A1:E10 数据，算一下总和

> 创建一个柱状图，X轴用 A 列，Y轴用 B 列
```

### 工具列表

| 工具 | 用途 |
|------|------|
| `excel_open(path)` | 打开/创建工作簿 |
| `excel_save(path?)` | 保存 |
| `excel_inspect()` | 查看工作簿结构 |
| `excel_inspect_sheet(sheet)` | 读取 sheet 数据 |
| `excel_execute_code(code)` | 执行 Python/C# 代码（最灵活） |
| `excel_execute_actions(actions)` | 批量 JSON 操作 |

### execute_code 可用变量

| 变量 | 含义 |
|------|------|
| `app` | Excel.Application COM 对象 |
| `wb` | 当前 ActiveWorkbook |
| `ws` | 当前 ActiveSheet |
| `result` | 设置此变量返回数据给 Claude |

## 注意事项

- 索引从 1 开始（COM 约定）
- 颜色是 **BGR** 格式：红=`0x0000FF`，蓝=`0xFF0000`，黄=`0x00FFFF`
- Excel 需要在前台运行（server 会自动连接或启动）
- 大文件（>50MB）打开/保存需要几秒，属正常现象
- 避免逐行 COM 调用，优先使用 `Range("A1:Z1000").Value2 = data` 批量写入

## 故障排查

| 问题 | 解决 |
|------|------|
| Server 启动失败 | 检查 `pip install pywin32 mcp` |
| "No workbook open" | 先调用 `excel_open` |
| COM 断开 | Excel 可能被关闭了，重新打开 |
| 0x800AC472 错误 | Excel 处于编辑模式，按 Esc 退出后重试 |
| C# 编译失败 | Cells 访问需要 `((dynamic)ws.Cells[r,c])` |
