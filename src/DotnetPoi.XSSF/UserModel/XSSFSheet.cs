namespace DotnetPoi.XSSF.UserModel;

public sealed class XSSFSheet
{
    private readonly SortedDictionary<int, XSSFRow> _rows = new();
    private readonly XSSFWorkbook _workbook;
    private XSSFDrawing? _drawing;

    internal XSSFSheet(XSSFWorkbook workbook, string sheetName, int sheetIndex)
    {
        _workbook = workbook;
        SheetName = sheetName;
        SheetIndex = sheetIndex;
    }

    internal string SheetName { get; }

    internal int SheetIndex { get; }

    public XSSFWorkbook getWorkbook()
    {
        return _workbook;
    }

    public XSSFRow createRow(int rownum)
    {
        if (rownum < 0)
        {
            throw new ArgumentException("Row number must be non-negative.", nameof(rownum));
        }

        var row = new XSSFRow(this, rownum);
        _rows[rownum] = row;
        return row;
    }

    public XSSFRow? getRow(int rownum)
    {
        return _rows.TryGetValue(rownum, out var row) ? row : null;
    }

    public XSSFDrawing createDrawingPatriarch()
    {
        return _drawing ??= new XSSFDrawing(this, _workbook.GetNextDrawingIndex());
    }

    public int getLastRowNum()
    {
        return _rows.Count == 0 ? 0 : _rows.Keys.Max();
    }

    internal IReadOnlyCollection<XSSFRow> Rows => _rows.Values;

    internal XSSFDrawing? Drawing => _drawing;
}
