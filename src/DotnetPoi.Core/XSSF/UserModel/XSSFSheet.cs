using DotnetPoi.SS.UserModel;
using DotnetPoi.SS.Util;

namespace DotnetPoi.XSSF.UserModel;

public sealed class XSSFSheet : ISheet
{
    private readonly SortedDictionary<int, XSSFRow> _rows = new();
    private readonly List<CellRangeAddress> _mergedRegions = new();
    private readonly SortedDictionary<int, int> _columnWidths = new();
    private readonly List<XSSFHyperlink> _hyperlinks = new();
    private readonly List<XSSFDataValidation> _dataValidations = new();
    private readonly List<XSSFConditionalFormatting> _condFormatting = new();
    private readonly XSSFWorkbook _workbook;
    private XSSFDrawing? _drawing;

    // Page margins (inches, default OOXML values)
    public double PageMarginBottom { get; set; } = 0.75;
    public double PageMarginFooter { get; set; } = 0.3;
    public double PageMarginHeader { get; set; } = 0.3;
    public double PageMarginLeft { get; set; } = 0.7;
    public double PageMarginRight { get; set; } = 0.7;
    public double PageMarginTop { get; set; } = 0.75;

    // Page setup
    public string? PageOrientation { get; set; } // "portrait" or "landscape" (null = default)
    public int? PaperSize { get; set; } // 1 = Letter, 9 = A4, null = default
    public int? Scale { get; set; } // percentage, 10-400, null = default
    public int? FitToWidth { get; set; } // null = default
    public int? FitToHeight { get; set; } // null = default

    // Header/footer
    public string? HeaderCenter { get; set; }
    public string? FooterCenter { get; set; }

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

    public void addMergedRegion(CellRangeAddress region)
    {
        ArgumentNullException.ThrowIfNull(region);
        _mergedRegions.Add(region);
    }

    public IReadOnlyList<CellRangeAddress> getMergedRegions() => _mergedRegions;

    internal IReadOnlyList<CellRangeAddress> MergedRegions => _mergedRegions;

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

    internal IReadOnlyDictionary<int, int> ColumnWidths =>
        new SortedDictionary<int, int>(_columnWidths);

    internal IReadOnlyList<XSSFHyperlink> Hyperlinks => _hyperlinks;

    public IReadOnlyList<XSSFDataValidation> DataValidations => _dataValidations;

    internal void AddHyperlink(XSSFHyperlink link)
    {
        ArgumentNullException.ThrowIfNull(link);
        _hyperlinks.Add(link);
    }

    public void AddDataValidation(XSSFDataValidation validation)
    {
        ArgumentNullException.ThrowIfNull(validation);
        _dataValidations.Add(validation);
    }

    public IReadOnlyList<XSSFConditionalFormatting> ConditionalFormatting => _condFormatting;

    public void AddConditionalFormatting(XSSFConditionalFormatting cf)
    {
        ArgumentNullException.ThrowIfNull(cf);
        _condFormatting.Add(cf);
    }

    IRow ISheet.createRow(int rownum) => createRow(rownum);

    IRow? ISheet.getRow(int rownum) => getRow(rownum);

    IWorkbook ISheet.getWorkbook() => getWorkbook();
}
