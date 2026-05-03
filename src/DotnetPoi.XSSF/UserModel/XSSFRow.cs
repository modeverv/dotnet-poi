namespace DotnetPoi.XSSF.UserModel;

public sealed class XSSFRow
{
    private readonly SortedDictionary<int, XSSFCell> _cells = new();
    private readonly XSSFSheet _sheet;
    private readonly int _rowNum;

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

    internal IReadOnlyCollection<XSSFCell> Cells => _cells.Values;
}
