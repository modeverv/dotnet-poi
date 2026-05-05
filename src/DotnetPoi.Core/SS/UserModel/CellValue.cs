namespace DotnetPoi.SS.UserModel;

/// <summary>
/// Mimics the data view of a cell.
/// Ported from org.apache.poi.ss.usermodel.CellValue.
/// </summary>
public sealed class CellValue
{
    public static readonly CellValue TRUE = new(CellType.Boolean, 0.0, true, null, 0);
    public static readonly CellValue FALSE = new(CellType.Boolean, 0.0, false, null, 0);

    private readonly CellType _cellType;
    private readonly double _numberValue;
    private readonly bool _booleanValue;
    private readonly string? _textValue;
    private readonly int _errorCode;

    private CellValue(CellType cellType, double numberValue, bool booleanValue, string? textValue, int errorCode)
    {
        _cellType = cellType;
        _numberValue = numberValue;
        _booleanValue = booleanValue;
        _textValue = textValue;
        _errorCode = errorCode;
    }

    public CellValue(double numberValue)
        : this(CellType.Numeric, numberValue, false, null, 0)
    {
    }

    public CellValue(string stringValue)
        : this(CellType.String, 0.0, false, stringValue, 0)
    {
    }

    public static CellValue valueOf(bool booleanValue) => booleanValue ? TRUE : FALSE;

    public static CellValue getError(int errorCode) => new(CellType.Error, 0.0, false, null, errorCode);

    public bool getBooleanValue() => _booleanValue;

    public double getNumberValue() => _numberValue;

    public string? getStringValue() => _textValue;

    public CellType getCellType() => _cellType;

    public byte getErrorValue() => (byte)_errorCode;
}
