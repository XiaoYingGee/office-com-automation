using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace ExcelMcp
{
    class Program
    {
        static ExcelEngine _engine;

        static void Main(string[] args)
        {
            _engine = new ExcelEngine();
            Console.InputEncoding = Encoding.UTF8;
            Console.OutputEncoding = Encoding.UTF8;

            while (true)
            {
                string message = ReadMessage();
                if (message == null) break;

                try
                {
                    JObject request = JObject.Parse(message);
                    JObject response = HandleRequest(request);
                    if (response != null)
                        WriteMessage(response.ToString(Formatting.None));
                }
                catch (Exception ex)
                {
                    JObject errResp = MakeError(-32603, ex.Message, null);
                    WriteMessage(errResp.ToString(Formatting.None));
                }
            }

            _engine.Close();
        }

        // ---- MCP stdio transport ----

        static string ReadMessage()
        {
            int contentLength = -1;
            string line;
            while ((line = Console.ReadLine()) != null)
            {
                if (line.StartsWith("Content-Length:"))
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
            string header = "Content-Length: " + bytes.Length + "\r\n\r\n";
            Console.Write(header);
            Console.Write(json);
            Console.Out.Flush();
        }

        // ---- JSON-RPC dispatch ----

        static JObject HandleRequest(JObject request)
        {
            string method = request["method"] != null ? (string)request["method"] : "";
            JToken id = request["id"];
            JObject pars = request["params"] as JObject;
            if (pars == null) pars = new JObject();

            switch (method)
            {
                case "initialize":
                    return MakeResult(id, GetInitializeResult());
                case "initialized":
                    return null;
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
            var r = new JObject();
            r.Add("protocolVersion", "2024-11-05");
            var caps = new JObject();
            caps.Add("tools", new JObject());
            r.Add("capabilities", caps);
            var info = new JObject();
            info.Add("name", "excel-mcp");
            info.Add("version", "1.0.0");
            r.Add("serverInfo", info);
            return r;
        }

        static JObject GetToolsList()
        {
            var tools = new JArray();
            tools.Add(MakeToolDef("excel_open",
                "Open or create an Excel workbook. Args: path (string), create (bool, optional).",
                MakeSchema(new string[] { "path" }, new string[] { "path", "create" })));
            tools.Add(MakeToolDef("excel_save",
                "Save the current workbook. Args: path (string, optional for SaveAs).",
                MakeSchema(new string[0], new string[] { "path" })));
            tools.Add(MakeToolDef("excel_inspect",
                "Return workbook structure: sheets with names, used ranges, row/col counts.",
                MakeSchema(new string[0], new string[0])));
            tools.Add(MakeToolDef("excel_inspect_sheet",
                "Return sheet data as 2D array. Args: sheet (name or index), max_rows, max_cols.",
                MakeSchema(new string[0], new string[] { "sheet", "max_rows", "max_cols" })));
            tools.Add(MakeToolDef("excel_execute_code",
                "Execute C# code in Excel process via CSharpCodeProvider. "
                + "Available: app (Excel.Application), wb (ActiveWorkbook), ws (ActiveSheet). "
                + "Return a string from your code. Colors are BGR. Indices are 1-based. "
                + "Example: ws.Range[\"A1\"].Value2 = \"Hello\"; return \"done\";",
                MakeSchema(new string[] { "code" }, new string[] { "code" })));
            tools.Add(MakeToolDef("excel_execute_actions",
                "Execute structured JSON actions (batch). "
                + "Actions: write_cell, read_cell, write_range, read_range, clear_range, merge_cells, "
                + "set_format, insert_rows, delete_rows, add_sheet, rename_sheet, delete_sheet, "
                + "sort_range, find_replace, calculate, export_pdf.",
                MakeSchema(new string[] { "actions" }, new string[] { "actions" })));

            var result = new JObject();
            result.Add("tools", tools);
            return result;
        }

        static JObject MakeSchema(string[] required, string[] properties)
        {
            var schema = new JObject();
            schema.Add("type", "object");
            var props = new JObject();
            foreach (string p in properties)
            {
                var prop = new JObject();
                prop.Add("type", "string");
                props.Add(p, prop);
            }
            schema.Add("properties", props);
            if (required.Length > 0)
                schema.Add("required", new JArray(required));
            return schema;
        }

        static JObject MakeToolDef(string name, string description, JObject inputSchema)
        {
            var tool = new JObject();
            tool.Add("name", name);
            tool.Add("description", description);
            tool.Add("inputSchema", inputSchema);
            return tool;
        }

        static JObject HandleToolCall(JObject pars)
        {
            string toolName = pars["name"] != null ? (string)pars["name"] : "";
            JObject arguments = pars["arguments"] as JObject;
            if (arguments == null) arguments = new JObject();

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
                        resultText = _engine.Inspect();
                        break;
                    case "excel_inspect_sheet":
                        resultText = ToolInspectSheet(arguments);
                        break;
                    case "excel_execute_code":
                        resultText = _engine.ExecuteCode((string)arguments["code"] ?? "");
                        break;
                    case "excel_execute_actions":
                        resultText = _engine.ExecuteActions((string)arguments["actions"] ?? "[]");
                        break;
                    default:
                        resultText = "{\"ok\":false,\"error\":\"Unknown tool: " + toolName + "\"}";
                        isError = true;
                        break;
                }
            }
            catch (Exception ex)
            {
                var err = new Dictionary<string, object>();
                err.Add("ok", false);
                err.Add("error", ex.Message);
                resultText = JsonConvert.SerializeObject(err);
                isError = true;
            }

            var content = new JArray();
            var textItem = new JObject();
            textItem.Add("type", "text");
            textItem.Add("text", resultText);
            content.Add(textItem);

            var response = new JObject();
            response.Add("content", content);
            response.Add("isError", isError);
            return response;
        }

        static string ToolOpen(JObject args)
        {
            string path = (string)args["path"];
            bool create = args["create"] != null && (bool)args["create"];
            if (create)
                _engine.Create(path);
            else
                _engine.Open(path);
            return _engine.Inspect();
        }

        static string ToolSave(JObject args)
        {
            string path = args["path"] != null ? (string)args["path"] : null;
            _engine.Save(path);
            return "{\"ok\":true}";
        }

        static string ToolInspectSheet(JObject args)
        {
            string sheet = args["sheet"] != null ? (string)args["sheet"] : "1";
            int maxRows = args["max_rows"] != null ? args["max_rows"].Value<int>() : 50;
            int maxCols = args["max_cols"] != null ? args["max_cols"].Value<int>() : 26;
            return _engine.InspectSheet(sheet, maxRows, maxCols);
        }

        // ---- Helpers ----

        static JObject MakeResult(JToken id, JObject result)
        {
            if (result == null) return null;
            var r = new JObject();
            r.Add("jsonrpc", "2.0");
            r.Add("id", id);
            r.Add("result", result);
            return r;
        }

        static JObject MakeError(int code, string message, JToken id)
        {
            var err = new JObject();
            err.Add("code", code);
            err.Add("message", message);
            var r = new JObject();
            r.Add("jsonrpc", "2.0");
            r.Add("id", id);
            r.Add("error", err);
            return r;
        }
    }
}
