using DotnetPoi.SS.UserModel;

namespace DotnetPoi.HSSF.UserModel;

public sealed class HSSFRow : IRow
{
    private readonly SortedDictionary<int, HSSFCell> _cells = new();
    private readonly HSSFSheet _sheet;
    private readonly int _rowNum;

    internal HSSFRow(HSSFSheet sheet, int rowNum)
    {
        _sheet = sheet;
        _rowNum = rowNum;
    }

    public HSSFCell createCell(int columnIndex)
    {
        if (columnIndex < 0)
        {
            throw new ArgumentException("Column index must be non-negative.", nameof(columnIndex));
        }

        var cell = new HSSFCell(this, columnIndex);
        _cells[columnIndex] = cell;
        return cell;
    }

    public HSSFCell? getCell(int cellnum) => _cells.TryGetValue(cellnum, out var cell) ? cell : null;

    public short getLastCellNum() => _cells.Count == 0 ? (short)-1 : (short)(_cells.Keys.Max() + 1);

    public int getRowNum() => _rowNum;

    public HSSFSheet getSheet() => _sheet;

    public void setHeight(float height)
    {
        // HSSF row height not yet implemented
    }

    public float getHeight() => 15.0f;

    public void setHidden(bool hidden) { }
    public bool isHidden() => false;

    internal IReadOnlyCollection<HSSFCell> Cells => _cells.Values;

    ICell IRow.createCell(int columnIndex) => createCell(columnIndex);

    ICell? IRow.getCell(int cellnum) => getCell(cellnum);

    ISheet IRow.getSheet() => getSheet();
}
