using DotnetPoi.SS.UserModel;

namespace DotnetPoi.HSSF.UserModel;

public sealed class HSSFCell : ICell
{
    private readonly HSSFRow _row;
    private readonly int _cellNum;
    private CellType _cellType = CellType.Blank;
    private string? _stringValue;
    private double _numericValue;
    private bool _booleanValue;
    private byte _errorValue;
    private HSSFCellStyle? _cellStyle;

    internal HSSFCell(HSSFRow row, int cellNum)
    {
        _row = row;
        _cellNum = cellNum;
    }

    public int getColumnIndex() => _cellNum;

    public int getRowIndex() => _row.getRowNum();

    public CellType getCellType() => _cellType;

    public CellType getCachedFormulaResultType()
    {
        if (_cellType != CellType.Formula)
        {
            throw new InvalidOperationException("Only formula cells have cached results.");
        }

        return CellType.Numeric;
    }

    public string? getCellFormula()
    {
        if (_cellType != CellType.Formula)
        {
            throw new InvalidOperationException($"Cannot get a formula value from a {_cellType} cell.");
        }

        return null;
    }

    public void setCellValue(string? value)
    {
        _stringValue = value ?? string.Empty;
        _cellType = CellType.String;
    }

    public void setCellValue(double value)
    {
        _numericValue = value;
        _cellType = CellType.Numeric;
    }

    public void setCellValue(bool value)
    {
        _booleanValue = value;
        _cellType = CellType.Boolean;
    }

    public void setCellFormula(string? formula)
    {
        if (formula is null)
        {
            _cellType = CellType.Blank;
            return;
        }

        // TODO: [dotnet-poi] Not yet ported
        // Original: poi/poi/src/main/java/org/apache/poi/hssf/usermodel/HSSFCell.java#setCellFormula
        // Reason: BIFF formula tokenization is deferred beyond this Phase 6 xls bootstrap slice.
        // Issue: Phase 6 HSSF formula write backlog
        throw new NotImplementedException("HSSFCell.SetFormula is not yet ported. See Phase 6 HSSF formula write backlog.");
    }

    public string getStringCellValue()
    {
        if (_cellType != CellType.String)
        {
            throw new InvalidOperationException($"Cannot get a string value from a {_cellType} cell.");
        }

        return _stringValue ?? string.Empty;
    }

    public double getNumericCellValue()
    {
        if (_cellType != CellType.Numeric)
        {
            throw new InvalidOperationException($"Cannot get a numeric value from a {_cellType} cell.");
        }

        return _numericValue;
    }

    public bool getBooleanCellValue()
    {
        if (_cellType != CellType.Boolean)
        {
            throw new InvalidOperationException($"Cannot get a boolean value from a {_cellType} cell.");
        }

        return _booleanValue;
    }

    public string getErrorCellString()
    {
        if (_cellType != CellType.Error)
        {
            throw new InvalidOperationException($"Cannot get an error value from a {_cellType} cell.");
        }

        return _errorValue switch
        {
            0x00 => "#NULL!",
            0x07 => "#DIV/0!",
            0x0F => "#VALUE!",
            0x17 => "#REF!",
            0x1D => "#NAME?",
            0x24 => "#NUM!",
            0x2A => "#N/A",
            _ => string.Empty
        };
    }

    public HSSFCellStyle getCellStyle() => _cellStyle ?? getSheet().getWorkbook().getCellStyleAt(0);

    public void setCellStyle(HSSFCellStyle? style) => _cellStyle = style;

    public void setCellStyle(ICellStyle? style) => setCellStyle(style as HSSFCellStyle);

    public HSSFSheet getSheet() => _row.getSheet();

    public HSSFRow getRow() => _row;

    internal void SetBlank() => _cellType = CellType.Blank;

    internal void SetError(byte value)
    {
        _errorValue = value;
        _cellType = CellType.Error;
    }

    ISheet ICell.getSheet() => getSheet();

    IRow ICell.getRow() => getRow();

    ICellStyle ICell.getCellStyle() => getCellStyle();

    public void setCachedFormulaResult(CellValue value)
    {
        // HSSF formula evaluation is not yet ported.
    }

    public IHyperlink? getHyperlink() => null;

    public void setHyperlink(IHyperlink? link)
    {
        // HSSF hyperlinks not yet implemented
    }
}
