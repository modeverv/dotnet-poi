using DotnetPoi.SS.UserModel;

namespace DotnetPoi.HSSF.UserModel;

public sealed class HSSFRow : IRow
{
    private readonly SortedDictionary<int, HSSFCell> _cells = new();
    private readonly HSSFSheet _sheet;
    private readonly int _rowNum;

    private short _heightTwips; // 0 = use default
    private bool _hidden;

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

    public void setHeight(float heightPoints)
    {
        _heightTwips = heightPoints <= 0 ? (short)0 : (short)(heightPoints * 20);
    }

    public float getHeight() => _heightTwips <= 0 ? 15.0f : (float)_heightTwips / 20.0f;

    public void setHidden(bool hidden) => _hidden = hidden;
    public bool isHidden() => _hidden;

    internal short HeightTwips => _heightTwips;
    internal bool Hidden => _hidden;

    internal IReadOnlyCollection<HSSFCell> Cells => _cells.Values;

    ICell IRow.createCell(int columnIndex) => createCell(columnIndex);

    ICell? IRow.getCell(int cellnum) => getCell(cellnum);

    ISheet IRow.getSheet() => getSheet();
}
