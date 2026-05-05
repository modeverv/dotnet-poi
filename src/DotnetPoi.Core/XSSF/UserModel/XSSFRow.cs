using DotnetPoi.SS.UserModel;

namespace DotnetPoi.XSSF.UserModel;

public sealed class XSSFRow : IRow
{
    private readonly SortedDictionary<int, XSSFCell> _cells = new();
    private readonly XSSFSheet _sheet;
    private readonly int _rowNum;
    private float _height = -1; // -1 = default height

    internal XSSFRow(XSSFSheet sheet, int rowNum)
    {
        _sheet = sheet;
        _rowNum = rowNum;
    }

    public XSSFSheet getSheet()
    {
        return _sheet;
    }

    public XSSFCell createCell(int columnIndex)
    {
        if (columnIndex < 0)
        {
            throw new ArgumentException("Column index must be non-negative.", nameof(columnIndex));
        }

        var cell = new XSSFCell(this, columnIndex);
        _cells[columnIndex] = cell;
        return cell;
    }

    public XSSFCell? getCell(int cellnum)
    {
        return _cells.TryGetValue(cellnum, out var cell) ? cell : null;
    }

    public short getLastCellNum()
    {
        return _cells.Count == 0 ? (short)-1 : (short)(_cells.Keys.Max() + 1);
    }

    public int getRowNum()
    {
        return _rowNum;
    }

    public void setHeight(float height)
    {
        if (height < 0)
            throw new ArgumentException("Row height must be non-negative.", nameof(height));
        _height = height;
    }

    public float getHeight() => _height >= 0 ? _height : 15.0f; // 15.0 = default

    /// <summary>True if a custom height was explicitly set.</summary>
    internal bool HasCustomHeight => _height >= 0;

    internal float HeightValue => _height;

    internal IReadOnlyCollection<XSSFCell> Cells => _cells.Values;

    ICell IRow.createCell(int columnIndex) => createCell(columnIndex);

    ICell? IRow.getCell(int cellnum) => getCell(cellnum);

    ISheet IRow.getSheet() => getSheet();
}
