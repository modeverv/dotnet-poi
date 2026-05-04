using DotnetPoi.SS.UserModel;

namespace DotnetPoi.HSSF.UserModel;

public sealed class HSSFSheet : ISheet
{
    private readonly SortedDictionary<int, HSSFRow> _rows = new();
    private readonly HSSFWorkbook _workbook;

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

    IRow ISheet.createRow(int rownum) => createRow(rownum);

    IRow? ISheet.getRow(int rownum) => getRow(rownum);

    IWorkbook ISheet.getWorkbook() => getWorkbook();
}
