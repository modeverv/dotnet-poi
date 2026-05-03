using System.Globalization;
using DotnetPoi.SS.UserModel;

namespace DotnetPoi.XSSF.UserModel;

public sealed class XSSFCell : ICell
{
    private readonly XSSFRow _row;
    private readonly int _cellNum;
    private CellType _cellType = CellType.Blank;
    private string? _stringValue;
    private double _numericValue;
    private XSSFCellStyle? _cellStyle;

    internal XSSFCell(XSSFRow row, int cellNum)
    {
        _row = row;
        _cellNum = cellNum;
    }

    public XSSFSheet getSheet()
    {
        return getRow().getSheet();
    }

    public XSSFRow getRow()
    {
        return _row;
    }

    public int getRowIndex()
    {
        return _row.getRowNum();
    }

    public int getColumnIndex()
    {
        return _cellNum;
    }

    public CellType getCellType()
    {
        return _cellType;
    }

    public void setCellValue(string? value)
    {
        _stringValue = value ?? string.Empty;
        _cellType = CellType.String;
    }

    public void setCellValue(double value)
    {
        _numericValue = value;
        _stringValue = null;
        _cellType = CellType.Numeric;
    }

    public string getStringCellValue()
    {
        if (_cellType != CellType.String)
        {
            throw new InvalidOperationException("Cannot get a string value from a non-string cell.");
        }

        return _stringValue ?? string.Empty;
    }

    public double getNumericCellValue()
    {
        if (_cellType != CellType.Numeric)
        {
            throw new InvalidOperationException("Cannot get a numeric value from a non-numeric cell.");
        }

        return _numericValue;
    }

    public XSSFCellStyle getCellStyle()
    {
        return _cellStyle ?? getSheet().getWorkbook().getCellStyleAt(0);
    }

    public void setCellStyle(XSSFCellStyle? style)
    {
        if (style is not null && !ReferenceEquals(style.Workbook, getSheet().getWorkbook()))
        {
            throw new ArgumentException("This Style does not belong to the supplied Workbook Styles Source. Are you trying to assign a style from one workbook to the cell of a different workbook?", nameof(style));
        }

        _cellStyle = style;
    }

    public void setCellStyle(ICellStyle? style)
    {
        setCellStyle(style as XSSFCellStyle);
    }

    internal string GetNumericText()
    {
        return _numericValue.ToString("G15", CultureInfo.InvariantCulture);
    }

    internal void SetCellStyleIndex(int styleIndex)
    {
        _cellStyle = getSheet().getWorkbook().getCellStyleAt(styleIndex);
    }

    ICellStyle ICell.getCellStyle() => getCellStyle();

    ISheet ICell.getSheet() => getSheet();

    IRow ICell.getRow() => getRow();
}
