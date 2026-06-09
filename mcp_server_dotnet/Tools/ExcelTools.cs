using System.ComponentModel;
using System.Text.Json;
using ModelContextProtocol.Server;

namespace ExcelMcpDotnet.Tools;

[McpServerToolType]
public static class ExcelTools
{
    private static readonly ExcelEngine _engine = new();

    [McpServerTool(Name = "excel_open"), Description(
        "Open or create an Excel workbook.\n\n" +
        "Args:\n" +
        "    path: File path (.xlsx). Relative paths resolved from CWD.\n" +
        "    create: If \"true\", create a new workbook at path. Otherwise open existing.\n\n" +
        "Returns: JSON {\"ok\": true, \"path\": \"...\", \"sheets\": [...], \"elapsed_ms\": ...}")]
    public static string Open(string path, string create = "false")
    {
        if (create == "true")
            return _engine.Create(path);
        return _engine.Open(path);
    }

    [McpServerTool(Name = "excel_save"), Description(
        "Save the current workbook. Optionally save-as to a new path.\n\n" +
        "Args:\n" +
        "    path: If provided, SaveAs to this path. Otherwise save in-place.\n\n" +
        "Returns: JSON {\"ok\": true, \"elapsed_ms\": ...}")]
    public static string Save(string? path = null)
    {
        return _engine.Save(path);
    }

    [McpServerTool(Name = "excel_inspect"), Description(
        "Return workbook structure: sheets with names, used ranges, row/col counts.\n\n" +
        "Returns: JSON {\"sheets\": [{\"name\": ..., \"used_range\": ..., \"rows\": ..., \"cols\": ...}]}")]
    public static string Inspect()
    {
        return _engine.Inspect();
    }

    [McpServerTool(Name = "excel_inspect_sheet"), Description(
        "Return sheet data as 2D array.\n\n" +
        "Args:\n" +
        "    sheet: Sheet name or 1-based index (as string). Default \"1\".\n" +
        "    max_rows: Maximum rows to return. Default \"50\".\n" +
        "    max_cols: Maximum columns to return. Default \"26\".")]
    public static string InspectSheet(string sheet = "1", string max_rows = "50", string max_cols = "26")
    {
        int maxR = int.TryParse(max_rows, out var r) ? r : 50;
        int maxC = int.TryParse(max_cols, out var c) ? c : 26;
        return _engine.InspectSheet(sheet, maxR, maxC);
    }

    [McpServerTool(Name = "excel_execute_code"), Description(
        "Execute C# code in Excel process via Roslyn compilation.\n\n" +
        "Available: app (Excel.Application), wb (ActiveWorkbook), ws (ActiveSheet).\n" +
        "Return a string from your code. Colors are BGR. Indices are 1-based.\n\n" +
        "Example:\n" +
        "    ws.Range[\"A1\"].Value2 = \"Hello\";\n" +
        "    ws.Range[\"A1\"].Font.Bold = true;\n" +
        "    return \"done\";\n\n" +
        "Example:\n" +
        "    // Bulk write with 2D array\n" +
        "    var data = new object[100, 2];\n" +
        "    for (int i = 0; i < 100; i++) { data[i,0] = i+1; data[i,1] = i*i; }\n" +
        "    ws.Range[\"A1:B100\"].Value2 = data;\n" +
        "    return \"wrote 100 rows\";\n\n" +
        "Returns: JSON {\"ok\": true, \"result\": ..., \"elapsed_ms\": ...}")]
    public static string ExecuteCode(string code)
    {
        return _engine.ExecuteCode(code);
    }

    [McpServerTool(Name = "excel_execute_actions"), Description(
        "Execute structured JSON actions (batch).\n\n" +
        "Actions: write_cell, read_cell, write_range, read_range, clear_range,\n" +
        "merge_cells, set_format, add_sheet, delete_sheet.\n\n" +
        "Example: [{\"action\":\"write_cell\",\"sheet\":\"Sheet1\",\"cell\":\"A1\",\"value\":42}]\n\n" +
        "Returns: JSON array of results [{\"ok\": true, \"action\": ...}, ...]")]
    public static string ExecuteActions(string actions)
    {
        return _engine.ExecuteActions(actions);
    }
}
