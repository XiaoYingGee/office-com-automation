# Excel MCP Server — .NET Framework 4.8 (Legacy)

## 为什么手写 MCP 协议？

本项目**刻意不使用**官方 C# MCP SDK（`ModelContextProtocol` NuGet），原因：

| 约束 | 说明 |
|------|------|
| **无需 .NET SDK** | 用 Windows 自带 `csc.exe` (C# 5) 编译，零安装依赖 |
| **CSharpCodeProvider** | 动态编译用户代码的 API，只在 .NET Framework 下可用 |
| **Wine 兼容** | .NET Framework 4.8 可在 Wine 下运行，.NET 8 不行 |
| **体积极小** | 编译产物仅 ExcelMcp.exe + Newtonsoft.Json.dll (~700KB) |

官方 SDK 需要 .NET 8+，与以上约束冲突。如需使用官方 SDK 版本，见 `../mcp_server_dotnet/`。

## 架构

```
Claude Code ──stdio JSON-RPC──→ Program.cs (手写协议) ──→ ExcelEngine.cs ──COM──→ Excel.exe
                                     │
                                     ├─ initialize / tools/list / tools/call
                                     └─ 6 个 tool: open, save, inspect, inspect_sheet,
                                                   execute_code, execute_actions
```

## 构建

```powershell
# 一键构建（自动下载 Newtonsoft.Json，用 csc.exe 编译）
.\build.ps1
```

输出：`ExcelMcp/bin/Release/net48/ExcelMcp.exe`

## 接入 Claude Code

```bash
claude mcp add excel-csharp -- D:\Workspace\AI\office-com-automation\mcp_server_csharp\ExcelMcp\bin\Release\net48\ExcelMcp.exe
```

## 协议实现细节

手写的 JSON-RPC 仅 ~80 行代码（Program.cs），处理：
- `initialize` → 返回 protocolVersion + capabilities
- `tools/list` → 返回 6 个 tool schema
- `tools/call` → dispatch 到 ExcelEngine 方法
- `shutdown` → 关闭 Excel COM

传输层：逐行读写 JSON (`Console.ReadLine` / `Console.WriteLine`)，符合 MCP stdio transport 规范。
