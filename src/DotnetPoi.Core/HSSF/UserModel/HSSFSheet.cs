using DotnetPoi.SS.UserModel;
using DotnetPoi.SS.Util;

namespace DotnetPoi.HSSF.UserModel;

public sealed class HSSFSheet : ISheet
{
    private readonly SortedDictionary<int, HSSFRow> _rows = new();
    private readonly HSSFWorkbook _workbook;
    private readonly List<CellRangeAddress> _mergedRegions = new();
    private readonly SortedDictionary<int, int> _columnWidths = new();

    internal HSSFSheet(HSSFWorkbook workbook, string sheetName)
    {
        _workbook = workbook;
        SheetName = sheetName;
    }

    public string SheetName { get; }

    public HSSFRow createRow(int rownum)
    {
        if (rownum < 0)
        {
            throw new ArgumentException("Row number must be non-negative.", nameof(rownum));
        }

        var row = new HSSFRow(this, rownum);
        _rows[rownum] = row;
        return row;
    }

    public HSSFRow? getRow(int rownum) => _rows.TryGetValue(rownum, out var row) ? row : null;

    public int getLastRowNum() => _rows.Count == 0 ? 0 : _rows.Keys.Max();

    public HSSFWorkbook getWorkbook() => _workbook;

    internal IReadOnlyCollection<HSSFRow> Rows => _rows.Values;

    public void addMergedRegion(CellRangeAddress region)
    {
        ArgumentNullException.ThrowIfNull(region);
        _mergedRegions.Add(region);
    }

    public IReadOnlyList<CellRangeAddress> getMergedRegions() => _mergedRegions;

    public void setColumnWidth(int columnIndex, int width)
    {
        if (width < 0)
            throw new ArgumentException("Column width must be non-negative.", nameof(width));
        _columnWidths[columnIndex] = width;
    }

    public int getColumnWidth(int columnIndex)
    {
        return _columnWidths.TryGetValue(columnIndex, out var width) ? width : 0;
    }

    public void createFreezePane(int colSplit, int rowSplit) { }
    public void createFreezePane(int colSplit, int rowSplit, int leftmostColumn, int topRow) { }
    public void setColumnHidden(int columnIndex, bool hidden) { }
    public bool isColumnHidden(int columnIndex) => false;

    public void protectSheet(bool protect) { }
    public bool isSheetProtected() => false;

    public void setAutoFilter(CellRangeAddress range) { }
    public CellRangeAddress? getAutoFilter() => null;

    IRow ISheet.createRow(int rownum) => createRow(rownum);

    IRow? ISheet.getRow(int rownum) => getRow(rownum);

    IWorkbook ISheet.getWorkbook() => getWorkbook();
}
