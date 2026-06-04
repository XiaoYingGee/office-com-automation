using System.Text.Json;
using System.Text.Json.Serialization;

namespace ExcelOps;

// ---------------------------------------------------------------------------
// OpRequest
// ---------------------------------------------------------------------------

public sealed class OpTarget
{
    [JsonPropertyName("sheet")]
    public string? Sheet { get; set; }

    [JsonPropertyName("range")]
    public string? Range { get; set; }
}

public sealed class SaveAs
{
    [JsonPropertyName("path")]
    public string Path { get; set; } = "";

    [JsonPropertyName("format")]
    public string Format { get; set; } = "xlsx";
}

public sealed class OpRequest
{
    [JsonPropertyName("op")]
    public string Op { get; set; } = "";

    [JsonPropertyName("path")]
    public string Path { get; set; } = "";

    [JsonPropertyName("target")]
    public OpTarget? Target { get; set; }

    [JsonPropertyName("params")]
    public JsonElement Params { get; set; }

    [JsonPropertyName("save_as")]
    public SaveAs? SaveAs { get; set; }
}

// ---------------------------------------------------------------------------
// OpResponse
// ---------------------------------------------------------------------------

public sealed class ExcelErrorDto
{
    [JsonPropertyName("category")]
    public string Category { get; set; } = "";

    [JsonPropertyName("code")]
    public int Code { get; set; }

    [JsonPropertyName("message")]
    public string Message { get; set; } = "";

    [JsonPropertyName("hint")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Hint { get; set; }
}

public sealed class OpResponse
{
    [JsonPropertyName("ok")]
    public bool Ok { get; set; }

    [JsonPropertyName("result")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public object? Result { get; set; }

    [JsonPropertyName("error")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public ExcelErrorDto? Error { get; set; }

    public static OpResponse Success(object result) =>
        new() { Ok = true, Result = result };

    public static OpResponse Failure(ExcelErrorDto error) =>
        new() { Ok = false, Error = error };
}

// ---------------------------------------------------------------------------
// Shared JsonSerializerOptions
// ---------------------------------------------------------------------------

public static class JsonOptions
{
    public static readonly JsonSerializerOptions Default = new()
    {
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
        // Enum is serialized as string via ExcelErrorDto (we use string property, not enum directly)
        WriteIndented = false,
    };
}
