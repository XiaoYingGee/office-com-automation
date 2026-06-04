namespace ExcelOps;

public enum ErrorCategory
{
    FileNotFound,
    FileLocked,
    InvalidArg,
    RangeParseError,
    SheetNotFound,
    UnsupportedFormat,
    ComError,
    MacroTrustDisabled,
    Timeout,
    Unknown,
}

public sealed class OpException : Exception
{
    public ErrorCategory Category { get; }
    public int Code { get; }
    public string? Hint { get; }

    public OpException(ErrorCategory category, string message, int code = 0, string? hint = null)
        : base(message)
    {
        Category = category;
        Code = code;
        Hint = hint;
    }
}
