# Excel MCP Server — .NET 8+ (Official SDK)

使用 [官方 C# MCP SDK](https://github.com/modelcontextprotocol/csharp-sdk) 实现的 Excel 自动化 MCP Server。

## 与 .NET Framework 版本的区别

| | mcp_server_csharp (Legacy) | mcp_server_dotnet (本目录) |
|---|---|---|
| Runtime | .NET Framework 4.8 | .NET 8+ |
| MCP 协议 | 手写 JSON-RPC | 官方 `ModelContextProtocol` SDK |
| 代码执行 | `CSharpCodeProvider` | Roslyn (`Microsoft.CodeAnalysis`) |
| Tool 定义 | 手写 schema | `[McpServerTool]` attribute |
| 编译 | `csc.exe`（零依赖） | `dotnet build`（需 .NET SDK） |
| Wine 兼容 | ✅ | ❌ |

## 前提条件

- Windows 10/11
- [.NET 8 SDK](https://dotnet.microsoft.com/download/dotnet/8.0)
- Microsoft Excel 已安装

## 构建

```bash
dotnet build -c Release
```

## 接入 Claude Code

```bash
# 项目级
claude mcp add excel-dotnet -- dotnet run --project D:\Workspace\AI\office-com-automation\mcp_server_dotnet\ExcelMcpDotnet.csproj

# 全局
claude mcp add excel-dotnet --scope user -- dotnet run --project D:\Workspace\AI\office-com-automation\mcp_server_dotnet\ExcelMcpDotnet.csproj
```

## 工具列表

| 工具 | 说明 |
|------|------|
| `excel_open` | 打开/创建工作簿 |
| `excel_save` | 保存工作簿 |
| `excel_inspect` | 查看工作簿结构 |
| `excel_inspect_sheet` | 读取 sheet 数据 |
| `excel_execute_code` | 执行 C# 代码（Roslyn 动态编译） |
| `excel_execute_actions` | 批量 JSON 操作 |
