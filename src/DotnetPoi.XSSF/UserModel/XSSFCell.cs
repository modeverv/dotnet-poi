using System.Globalization;

namespace DotnetPoi.XSSF.UserModel;

public enum XSSFCellType
{
    Blank,
    Numeric,
    String
}

public sealed class XSSFCell
{
    private readonly XSSFRow _row;
    private readonly int _cellNum;
    private XSSFCellType _cellType = XSSFCellType.Blank;
    private string? _stringValue;
    private double _numericValue;

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

    public XSSFCellType getCellType()
    {
        return _cellType;
    }

    public void setCellValue(string? value)
    {
        _stringValue = value ?? string.Empty;
        _cellType = XSSFCellType.String;
    }

    public void setCellValue(double value)
    {
        _numericValue = value;
        _stringValue = null;
        _cellType = XSSFCellType.Numeric;
    }

    public string getStringCellValue()
    {
        if (_cellType != XSSFCellType.String)
        {
            throw new InvalidOperationException("Cannot get a string value from a non-string cell.");
        }

        return _stringValue ?? string.Empty;
    }

    public double getNumericCellValue()
    {
        if (_cellType != XSSFCellType.Numeric)
        {
            throw new InvalidOperationException("Cannot get a numeric value from a non-numeric cell.");
        }

        return _numericValue;
    }

    internal string GetNumericText()
    {
        return _numericValue.ToString("G15", CultureInfo.InvariantCulture);
    }
}
