using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using ModelContextProtocol.Server;

var builder = Host.CreateApplicationBuilder(args);

builder.Services.AddMcpServer(options =>
{
    options.ServerInfo = new()
    {
        Name = "excel-dotnet",
        Version = "1.0.0"
    };
    options.ServerInstructions =
        "Excel automation server (.NET 8 + official MCP SDK). " +
        "Use excel_inspect to see workbook structure, " +
        "excel_execute_code to run C# code in Excel process (most flexible), " +
        "or excel_execute_actions for structured JSON operations.";
})
.WithStdioServerTransport()
.WithToolsFromAssembly();

await builder.Build().RunAsync();
