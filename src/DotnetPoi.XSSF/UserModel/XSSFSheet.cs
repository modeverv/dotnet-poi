using DotnetPoi.SS.UserModel;

namespace DotnetPoi.XSSF.UserModel;

public sealed class XSSFSheet : ISheet
{
    private readonly SortedDictionary<int, XSSFRow> _rows = new();
    private readonly XSSFWorkbook _workbook;
    private XSSFDrawing? _drawing;

    internal XSSFSheet(XSSFWorkbook workbook, string sheetName, int sheetIndex, int sheetId = 0, bool isHidden = false)
    {
        _workbook = workbook;
        SheetName = sheetName;
        SheetIndex = sheetIndex;
        SheetId = sheetId != 0 ? sheetId : sheetIndex;
        IsHidden = isHidden;
    }

    internal string SheetName { get; }

    internal int SheetIndex { get; }

    internal int SheetId { get; }

    internal bool IsHidden { get; }

    public XSSFWorkbook getWorkbook()
    {
        return _workbook;
    }

    internal bool IsRowsDirty { get; private set; }

    // Preserved original worksheet XML bytes for round-trip fidelity
    internal byte[]? PreservedWorksheetXml { get; set; }

    public XSSFRow createRow(int rownum)
    {
        if (rownum < 0)
        {
            throw new ArgumentException("Row number must be non-negative.", nameof(rownum));
        }

        if (!_workbook.IsLoading) IsRowsDirty = true;
        var row = new XSSFRow(this, rownum);
        _rows[rownum] = row;
        _workbook.MarkDirty();
        return row;
    }

    public XSSFRow? getRow(int rownum)
    {
        return _rows.TryGetValue(rownum, out var row) ? row : null;
    }

    public XSSFDrawing createDrawingPatriarch()
    {
        _workbook.MarkDirty();
        return _drawing ??= new XSSFDrawing(this, _workbook.GetNextDrawingIndex());
    }

    public int getLastRowNum()
    {
        return _rows.Count == 0 ? 0 : _rows.Keys.Max();
    }

    internal IReadOnlyCollection<XSSFRow> Rows => _rows.Values;

    internal XSSFDrawing? Drawing => _drawing;

    // Preserved original bytes for drawing and its rels when loaded from file
    internal byte[]? PreservedDrawingXml { get; set; }
    internal byte[]? PreservedDrawingRelsXml { get; set; }

    IRow ISheet.createRow(int rownum) => createRow(rownum);

    IRow? ISheet.getRow(int rownum) => getRow(rownum);

    IWorkbook ISheet.getWorkbook() => getWorkbook();
}
