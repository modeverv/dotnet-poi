using DotnetPoi.SS.UserModel;

namespace DotnetPoi.XSSF.UserModel;

public sealed class XSSFDataFormat : IDataFormat
{
    public const int FIRST_USER_DEFINED_FORMAT_INDEX = 164;

    private static readonly string[] BuiltinFormats =
    {
        "General", "0", "0.00", "#,##0", "#,##0.00",
        "\"$\"#,##0_);(\"$\"#,##0)", "\"$\"#,##0_);[Red](\"$\"#,##0)",
        "\"$\"#,##0.00_);(\"$\"#,##0.00)", "\"$\"#,##0.00_);[Red](\"$\"#,##0.00)",
        "0%", "0.00%", "0.00E+00", "# ?/?", "# ??/??", "m/d/yy", "d-mmm-yy",
        "d-mmm", "mmm-yy", "h:mm AM/PM", "h:mm:ss AM/PM", "h:mm", "h:mm:ss",
        "m/d/yy h:mm", "reserved-0x17", "reserved-0x18", "reserved-0x19",
        "reserved-0x1A", "reserved-0x1B", "reserved-0x1C", "reserved-0x1D",
        "reserved-0x1E", "reserved-0x1F", "reserved-0x20", "reserved-0x21",
        "reserved-0x22", "reserved-0x23", "reserved-0x24", "#,##0_);(#,##0)",
        "#,##0_);[Red](#,##0)", "#,##0.00_);(#,##0.00)",
        "#,##0.00_);[Red](#,##0.00)", "_(* #,##0_);_(* (#,##0);_(* \"-\"_);_(@_)",
        "_(\"$\"* #,##0_);_(\"$\"* (#,##0);_(\"$\"* \"-\"_);_(@_)",
        "_(* #,##0.00_);_(* (#,##0.00);_(* \"-\"??_);_(@_)",
        "_(\"$\"* #,##0.00_);_(\"$\"* (#,##0.00);_(\"$\"* \"-\"??_);_(@_)",
        "mm:ss", "[h]:mm:ss", "mm:ss.0", "##0.0E+0", "@"
    };

    private readonly XSSFWorkbook _workbook;

    internal XSSFDataFormat(XSSFWorkbook workbook)
    {
        _workbook = workbook;
    }

    public short getFormat(string format)
    {
        ArgumentNullException.ThrowIfNull(format);
        var normalizedFormat = string.Equals(format, "text", StringComparison.OrdinalIgnoreCase) ? "@" : format;
        for (var i = 0; i < BuiltinFormats.Length; i++)
        {
            if (BuiltinFormats[i] == normalizedFormat)
            {
                return (short)i;
            }
        }

        return (short)_workbook.GetOrAddCustomNumberFormat(normalizedFormat);
    }

    public string? getFormat(short index)
    {
        var unsignedIndex = index & 0xffff;
        if (unsignedIndex >= 0 && unsignedIndex < BuiltinFormats.Length)
        {
            return BuiltinFormats[unsignedIndex];
        }

        return _workbook.GetCustomNumberFormat(unsignedIndex);
    }
}
