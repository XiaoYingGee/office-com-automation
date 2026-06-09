using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace ExcelMcp
{
    /// <summary>
    /// MCP Server (stdio JSON-RPC 2.0) for Excel automation.
    ///
    /// Tools:
    ///   excel_open           — open/create workbook
    ///   excel_save           — save workbook
    ///   excel_inspect        — workbook structure
    ///   excel_inspect_sheet  — sheet data
    ///   excel_execute_code   — CodeAct: compile+run C# code in-process
    ///   excel_execute_actions— structured JSON actions
    ///
    /// Usage:
    ///   ExcelMcp.exe         (reads JSON-RPC from stdin, writes to stdout)
    /// </summary>
    class Program
    {
        static ExcelEngine _engine;

        static void Main(string[] args)
        {
            _engine = new ExcelEngine();
            Console.InputEncoding = Encoding.UTF8;
            Console.OutputEncoding = Encoding.UTF8;

            // MCP stdio transport: read Content-Length header + JSON body
            while (true)
            {
                string message = ReadMessage();
                if (message == null) break;

                try
                {
                    JObject request = JObject.Parse(message);
                    JObject response = HandleRequest(request);
                    WriteMessage(response.ToString(Formatting.None));
                }
                catch (Exception ex)
                {
                    var errResp = MakeError(-32603, ex.Message, null);
                    WriteMessage(errResp.ToString(Formatting.None));
                }
            }

            _engine.Close();
        }

        // ---- MCP stdio transport (Content-Length framing) ----

        static string ReadMessage()
        {
            int contentLength = -1;
            string line;
            while ((line = Console.ReadLine()) != null)
            {
                if (line.StartsWith("Content-Length:", StringComparison.OrdinalIgnoreCase))
                {
                    contentLength = int.Parse(line.Substring("Content-Length:".Length).Trim());
                }
                else if (line == "")
                {
                    break;
                }
            }

            if (contentLength < 0) return null;

            char[] buffer = new char[contentLength];
            int read = 0;
            while (read < contentLength)
            {
                int n = Console.In.Read(buffer, read, contentLength - read);
                if (n <= 0) return null;
                read += n;
            }
            return new string(buffer);
        }

        static void WriteMessage(string json)
        {
            byte[] bytes = Encoding.UTF8.GetBytes(json);
            string header = $"Content-Length: {bytes.Length}\r\n\r\n";
            Console.Write(header);
            Console.Write(json);
            Console.Out.Flush();
        }

        // ---- JSON-RPC dispatch ----

        static JObject HandleRequest(JObject request)
        {
            string method = (string)request["method"] ?? "";
            JToken id = request["id"];
            JObject pars = request["params"] as JObject ?? new JObject();

            switch (method)
            {
                case "initialize":
                    return MakeResult(id, GetInitializeResult());
                case "initialized":
                    return null; // notification, no response needed
                case "tools/list":
                    return MakeResult(id, GetToolsList());
                case "tools/call":
                    return MakeResult(id, HandleToolCall(pars));
                case "shutdown":
                    _engine.Close();
                    return MakeResult(id, new JObject());
                default:
                    return MakeError(-32601, "Method not found: " + method, id);
            }
        }

        static JObject GetInitializeResult()
        {
            return JObject.FromObject(new
            {
                protocolVersion = "2024-11-05",
                capabilities = new { tools = new { } },
                serverInfo = new { name = "excel-mcp", version = "1.0.0" }
            });
        }

        static JObject GetToolsList()
        {
            var tools = new JArray
            {
                MakeToolDef("excel_open",
                    "Open or create an Excel workbook. Args: path (string), create (bool, optional).",
                    new { type = "object", properties = new {
                        path = new { type = "string", description = "File path (.xlsx)" },
                        create = new { type = "boolean", description = "Create new if true" }
                    }, required = new[] { "path" } }),

                MakeToolDef("excel_save",
                    "Save the current workbook. Args: path (string, optional — SaveAs if provided).",
                    new { type = "object", properties = new {
                        path = new { type = "string", description = "SaveAs path (optional)" }
                    } }),

                MakeToolDef("excel_inspect",
                    "Return workbook structure: sheets with names, used ranges, row/col counts.",
                    new { type = "object", properties = new { } }),

                MakeToolDef("excel_inspect_sheet",
                    "Return sheet data as 2D array. Args: sheet (name or index), max_rows, max_cols.",
                    new { type = "object", properties = new {
                        sheet = new { type = "string", description = "Sheet name or 1-based index", @default = "1" },
                        max_rows = new { type = "integer", description = "Max rows to return", @default = 50 },
                        max_cols = new { type = "integer", description = "Max cols to return", @default = 26 }
                    } }),

                MakeToolDef("excel_execute_code",
                    @"Execute C# code in Excel process. The code is compiled via CSharpCodeProvider and run in-process.

Available variables (passed as parameters to your Execute method):
  Excel.Application app  — the running Excel instance
  Excel.Workbook wb      — ActiveWorkbook
  Excel.Worksheet ws     — ActiveSheet

Return a string (will be sent back as result). Use null for no return value.

Example:
  ws.Range[""A1""].Value2 = ""Hello"";
  ws.Range[""A1""].Font.Bold = true;
  return ""done"";

Example (read data):
  var used = ws.UsedRange;
  return $""{used.Rows.Count} rows x {used.Columns.Count} cols"";

Example (bulk write):
  for (int i = 1; i <= 100; i++) {
      ws.Cells[i, 1].Value2 = i;
      ws.Cells[i, 2].Value2 = i * i;
  }
  ws.Range[""A1:B100""].Columns.AutoFit();
  return ""wrote 100 rows"";

Note: Colors are BGR (red = 0x0000FF, blue = 0xFF0000). Indices are 1-based.",
                    new { type = "object", properties = new {
                        code = new { type = "string", description = "C# code body (inside a method returning string)" }
                    }, required = new[] { "code" } }),

                MakeToolDef("excel_execute_actions",
                    @"Execute structured JSON actions (batch). Each action: {""action"":""write_cell"",""sheet"":""Sheet1"",""cell"":""A1"",""value"":42}
Available actions: write_cell, read_cell, write_range, read_range, clear_range, merge_cells, unmerge_cells, set_format, set_border, insert_rows, delete_rows, insert_cols, delete_cols, add_sheet, rename_sheet, delete_sheet, sort_range, find_replace, calculate, export_pdf, and more.",
                    new { type = "object", properties = new {
                        actions = new { type = "string", description = "JSON string: single action object or array of actions" }
                    }, required = new[] { "actions" } })
            };

            return new JObject { ["tools"] = tools };
        }

        static JObject MakeToolDef(string name, string description, object inputSchema)
        {
            return new JObject
            {
                ["name"] = name,
                ["description"] = description,
                ["inputSchema"] = JObject.FromObject(inputSchema)
            };
        }

        static JObject HandleToolCall(JObject pars)
        {
            string toolName = (string)pars["name"] ?? "";
            JObject arguments = pars["arguments"] as JObject ?? new JObject();

            string resultText;
            bool isError = false;

            try
            {
                switch (toolName)
                {
                    case "excel_open":
                        resultText = ToolOpen(arguments);
                        break;
                    case "excel_save":
                        resultText = ToolSave(arguments);
                        break;
                    case "excel_inspect":
                        resultText = ToolInspect();
                        break;
                    case "excel_inspect_sheet":
                        resultText = ToolInspectSheet(arguments);
                        break;
                    case "excel_execute_code":
                        resultText = ToolExecuteCode(arguments);
                        break;
                    case "excel_execute_actions":
                        resultText = ToolExecuteActions(arguments);
                        break;
                    default:
                        resultText = JsonConvert.SerializeObject(new { ok = false, error = "Unknown tool: " + toolName });
                        isError = true;
                        break;
                }
            }
            catch (Exception ex)
            {
                resultText = JsonConvert.SerializeObject(new { ok = false, error = ex.Message, traceback = ex.StackTrace });
                isError = true;
            }

            return new JObject
            {
                ["content"] = new JArray { new JObject { ["type"] = "text", ["text"] = resultText } },
                ["isError"] = isError
            };
        }

        // ---- Tool implementations ----

        static string ToolOpen(JObject args)
        {
            string path = (string)args["path"];
            bool create = args["create"]?.Value<bool>() ?? false;
            if (create)
                _engine.Create(path);
            else
                _engine.Open(path);
            var info = _engine.Inspect();
            return JsonConvert.SerializeObject(new { ok = true, path = _engine.FilePath, sheets = info });
        }

        static string ToolSave(JObject args)
        {
            string path = (string)args["path"];
            _engine.Save(path);
            return JsonConvert.SerializeObject(new { ok = true });
        }

        static string ToolInspect()
        {
            return JsonConvert.SerializeObject(_engine.Inspect());
        }

        static string ToolInspectSheet(JObject args)
        {
            string sheet = (string)args["sheet"] ?? "1";
            int maxRows = args["max_rows"]?.Value<int>() ?? 50;
            int maxCols = args["max_cols"]?.Value<int>() ?? 26;
            return JsonConvert.SerializeObject(_engine.InspectSheet(sheet, maxRows, maxCols));
        }

        static string ToolExecuteCode(JObject args)
        {
            string code = (string)args["code"] ?? "";
            return _engine.ExecuteCode(code);
        }

        static string ToolExecuteActions(JObject args)
        {
            string actionsJson = (string)args["actions"] ?? "[]";
            return _engine.ExecuteActions(actionsJson);
        }

        // ---- Helpers ----

        static JObject MakeResult(JToken id, JObject result)
        {
            if (result == null) return null;
            return new JObject
            {
                ["jsonrpc"] = "2.0",
                ["id"] = id,
                ["result"] = result
            };
        }

        static JObject MakeError(int code, string message, JToken id)
        {
            return new JObject
            {
                ["jsonrpc"] = "2.0",
                ["id"] = id,
                ["error"] = new JObject { ["code"] = code, ["message"] = message }
            };
        }
    }
}
