using System.Globalization;
using DotnetPoi.SS.UserModel;

namespace DotnetPoi.XSSF.UserModel;

/// <summary>
/// Ported from org.apache.poi.xssf.usermodel.XSSFCell.
/// </summary>
public sealed class XSSFCell : ICell
{
    private readonly XSSFRow _row;
    private readonly int _cellNum;
    private CellType _cellType = CellType.Blank;

    // For FORMULA cells: the cached result type and whether this is a formula
    private bool _isFormula;
    private string? _formula;
    private CellType _cachedFormulaResultType = CellType.Numeric;
    private bool _hasCachedValue;

    // Typed value storage (matches POI's CTCell fields)
    private string? _stringValue;
    private double _numericValue;
    private bool _booleanValue;
    private string? _errorString; // raw OOXML error string e.g. "#DIV/0!"

    private XSSFCellStyle? _cellStyle;

    internal XSSFCell(XSSFRow row, int cellNum)
    {
        _row = row;
        _cellNum = cellNum;
    }

    public XSSFSheet getSheet() => _row.getSheet();
    public XSSFRow getRow() => _row;
    public int getRowIndex() => _row.getRowNum();
    public int getColumnIndex() => _cellNum;

    /// <summary>
    /// Returns FORMULA if the cell has a formula, otherwise the base type.
    /// Ported from XSSFCell.getCellType().
    /// </summary>
    public CellType getCellType() => _cellType;

    /// <summary>
    /// Only valid for formula cells.
    /// Ported from XSSFCell.getCachedFormulaResultType() / getBaseCellType().
    /// </summary>
    public CellType getCachedFormulaResultType()
    {
        if (_cellType != CellType.Formula)
            throw new InvalidOperationException("Only formula cells have cached results.");
        return _cachedFormulaResultType;
    }

    /// <summary>
    /// Returns the formula string for formula cells.
    /// Ported from XSSFCell.getCellFormula().
    /// </summary>
    public string? getCellFormula()
    {
        if (_cellType != CellType.Formula)
            throw new InvalidOperationException($"Cannot get a formula value from a {_cellType} cell.");
        return _formula;
    }

    // ----- setCellValue overloads -----

    public void setCellValue(string? value)
    {
        _stringValue = value ?? string.Empty;
        _hasCachedValue = true;
        if (_isFormula)
        {
            _cachedFormulaResultType = CellType.String;
        }
        else
        {
            _cellType = CellType.String;
        }
        _row.getSheet().getWorkbook().MarkDirty();
    }

    public void setCellValue(double value)
    {
        _numericValue = value;
        _stringValue = null;
        _hasCachedValue = true;
        if (_isFormula)
        {
            _cachedFormulaResultType = CellType.Numeric;
        }
        else
        {
            _cellType = CellType.Numeric;
        }
        _row.getSheet().getWorkbook().MarkDirty();
    }

    public void setCellValue(bool value)
    {
        _booleanValue = value;
        _hasCachedValue = true;
        if (_isFormula)
        {
            _cachedFormulaResultType = CellType.Boolean;
        }
        else
        {
            _cellType = CellType.Boolean;
        }
        _row.getSheet().getWorkbook().MarkDirty();
    }

    /// <summary>
    /// Sets the formula string without evaluating it.
    /// Ported from XSSFCell.setCellFormula(String).
    /// </summary>
    public void setCellFormula(string? formula)
    {
        if (formula is null)
        {
            _formula = null;
            _isFormula = false;
            _cellType = _hasCachedValue ? _cachedFormulaResultType : CellType.Blank;
            _row.getSheet().getWorkbook().MarkDirty();
            return;
        }

        _formula = formula;
        _isFormula = true;
        _cellType = CellType.Formula;
        _cachedFormulaResultType = _hasCachedValue ? _cachedFormulaResultType : CellType.Numeric;
        _row.getSheet().getWorkbook().MarkDirty();
    }

    // ----- getters -----

    /// <summary>
    /// Returns the string value. Valid for STRING and FORMULA cells with STRING cached type.
    /// Ported from XSSFCell.getStringCellValue().
    /// </summary>
    public string getStringCellValue()
    {
        var effectiveType = _isFormula ? _cachedFormulaResultType : _cellType;
        if (effectiveType != CellType.String)
            throw new InvalidOperationException($"Cannot get a string value from a {_cellType} cell.");
        return _stringValue ?? string.Empty;
    }

    /// <summary>
    /// Returns the numeric value. Valid for NUMERIC and FORMULA cells with NUMERIC cached type.
    /// Ported from XSSFCell.getNumericCellValue().
    /// </summary>
    public double getNumericCellValue()
    {
        var effectiveType = _isFormula ? _cachedFormulaResultType : _cellType;
        if (effectiveType != CellType.Numeric)
            throw new InvalidOperationException($"Cannot get a numeric value from a {_cellType} cell.");
        return _numericValue;
    }

    /// <summary>
    /// Returns the boolean value. Valid for BOOLEAN and FORMULA cells with BOOLEAN cached type.
    /// Ported from XSSFCell.getBooleanCellValue(): checks _cell.isSetV() &amp;&amp; "1".equals(_cell.getV()).
    /// </summary>
    public bool getBooleanCellValue()
    {
        var effectiveType = _isFormula ? _cachedFormulaResultType : _cellType;
        if (effectiveType != CellType.Boolean)
            throw new InvalidOperationException($"Cannot get a boolean value from a {_cellType} cell.");
        return _booleanValue;
    }

    /// <summary>
    /// Returns the raw OOXML error string (e.g. "#DIV/0!") for ERROR cells.
    /// Ported from XSSFCell.getErrorCellString().
    /// </summary>
    public string getErrorCellString()
    {
        var effectiveType = _isFormula ? _cachedFormulaResultType : _cellType;
        if (effectiveType != CellType.Error)
            throw new InvalidOperationException($"Cannot get an error value from a {_cellType} cell.");
        return _errorString ?? string.Empty;
    }

    public XSSFCellStyle getCellStyle() =>
        _cellStyle ?? getSheet().getWorkbook().getCellStyleAt(0);

    public void setCellStyle(XSSFCellStyle? style)
    {
        if (style is not null && !ReferenceEquals(style.Workbook, getSheet().getWorkbook()))
            throw new ArgumentException(
                "This Style does not belong to the supplied Workbook Styles Source.",
                nameof(style));
        _cellStyle = style;
    }

    public void setCellStyle(ICellStyle? style) => setCellStyle(style as XSSFCellStyle);

    // ----- internal helpers used by the reader -----

    /// <summary>
    /// Called by the reader when a formula cell is encountered.
    /// Sets FORMULA type and stores the cached value according to the t attribute.
    /// Ported from the reader side of XSSFCell.
    /// </summary>
    internal void SetFormulaWithCachedValue(CellType cachedType, string? rawValue)
    {
        SetFormulaFromXml(null, cachedType, rawValue, hasCachedValue: rawValue is not null);
    }

    internal void SetFormulaFromXml(string? formula, CellType cachedType, string? rawValue, bool hasCachedValue)
    {
        _isFormula = true;
        _cellType = CellType.Formula;
        _formula = formula;
        _cachedFormulaResultType = cachedType;
        _hasCachedValue = hasCachedValue;
        if (hasCachedValue)
            ApplyCachedValue(cachedType, rawValue);
    }

    internal void SetFormulaCachedValue(CellValue value)
    {
        if (_cellType != CellType.Formula)
            throw new InvalidOperationException("Only formula cells can receive cached formula values.");

        switch (value.getCellType())
        {
            case CellType.Numeric:
                setCellValue(value.getNumberValue());
                break;
            case CellType.String:
                setCellValue(value.getStringValue() ?? string.Empty);
                break;
            case CellType.Boolean:
                setCellValue(value.getBooleanValue());
                break;
            case CellType.Error:
                _cachedFormulaResultType = CellType.Error;
                _errorString = ErrorCodeToText(value.getErrorValue());
                _hasCachedValue = true;
                break;
            default:
                _cachedFormulaResultType = CellType.Blank;
                _hasCachedValue = false;
                break;
        }
    }

    /// <summary>
    /// Called by the reader for non-formula cells.
    /// Matches POI's getBaseCellType() logic with the t attribute.
    /// </summary>
    internal void SetValueFromXml(CellType baseCellType, string? rawValue)
    {
        _isFormula = false;
        _formula = null;
        _cellType = baseCellType;
        _hasCachedValue = rawValue is not null;
        ApplyCachedValue(baseCellType, rawValue);
    }

    private void ApplyCachedValue(CellType type, string? rawValue)
    {
        switch (type)
        {
            case CellType.Numeric:
                _numericValue = rawValue is not null
                    ? double.Parse(rawValue, CultureInfo.InvariantCulture)
                    : 0.0;
                break;
            case CellType.String:
                _stringValue = rawValue ?? string.Empty;
                break;
            case CellType.Boolean:
                // POI: "1" = true, "0" = false
                _booleanValue = rawValue == "1";
                break;
            case CellType.Error:
                _errorString = rawValue;
                break;
            case CellType.Blank:
                break;
        }
    }

    internal string GetNumericText() =>
        _numericValue.ToString("G15", CultureInfo.InvariantCulture);

    internal bool HasCachedValue => _hasCachedValue;

    private static string ErrorCodeToText(byte errorCode) =>
        errorCode switch
        {
            7 => "#DIV/0!",
            15 => "#VALUE!",
            23 => "#REF!",
            29 => "#NAME?",
            36 => "#NUM!",
            42 => "#N/A",
            _ => "#VALUE!"
        };

    internal void SetCellStyleIndex(int styleIndex) =>
        _cellStyle = getSheet().getWorkbook().getCellStyleAt(styleIndex);

    ICellStyle ICell.getCellStyle() => getCellStyle();
    ISheet ICell.getSheet() => getSheet();
    IRow ICell.getRow() => getRow();
}
