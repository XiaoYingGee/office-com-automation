using System.Text.Json;

namespace ExcelOps;

class Program
{
    static int Main(string[] args)
    {
        // Read ALL of stdin (capctl closes the pipe after writing the request)
        string input = Console.In.ReadToEnd();

        OpResponse response;
        try
        {
            var req = JsonSerializer.Deserialize<OpRequest>(input, JsonOptions.Default);
            if (req is null)
            {
                response = OpResponse.Failure(new ExcelErrorDto
                {
                    Category = nameof(ErrorCategory.InvalidArg),
                    Code = 0,
                    Message = "Failed to deserialize OpRequest: null result",
                });
            }
            else
            {
                var result = Ops.Dispatch(req);
                response = OpResponse.Success(result);
            }
        }
        catch (OpException ox)
        {
            response = OpResponse.Failure(new ExcelErrorDto
            {
                Category = ox.Category.ToString(),
                Code = ox.Code,
                Message = ox.Message,
                Hint = ox.Hint,
            });
        }
        catch (Exception ex)
        {
            response = OpResponse.Failure(new ExcelErrorDto
            {
                Category = nameof(ErrorCategory.Unknown),
                Code = 0,
                Message = ex.Message,
            });
        }

        string json = JsonSerializer.Serialize(response, JsonOptions.Default);
        Console.Write(json);

        // Always exit 0 — a structured error response was emitted
        return 0;
    }
}
