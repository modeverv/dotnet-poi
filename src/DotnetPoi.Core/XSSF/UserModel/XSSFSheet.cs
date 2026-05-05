using System.Globalization;
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
    private readonly List<XSSFPivotTable> _pivotTables = new();
    private readonly XSSFWorkbook _workbook;
    private XSSFDrawing? _drawing;

    // Freeze panes
    public int FreezeColSplit { get; private set; } = -1; // -1 = not set
    public int FreezeRowSplit { get; private set; } = -1;

    // Hidden columns
    private readonly HashSet<int> _hiddenColumns = new();

    // Sheet protection
    private bool _sheetProtected;

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

    public void createFreezePane(int colSplit, int rowSplit)
    {
        createFreezePane(colSplit, rowSplit, colSplit, rowSplit);
    }

    public void createFreezePane(int colSplit, int rowSplit, int leftmostColumn, int topRow)
    {
        FreezeColSplit = colSplit;
        FreezeRowSplit = rowSplit;
    }

    public void setColumnHidden(int columnIndex, bool hidden)
    {
        if (hidden)
            _hiddenColumns.Add(columnIndex);
        else
            _hiddenColumns.Remove(columnIndex);
    }

    public bool isColumnHidden(int columnIndex)
    {
        return _hiddenColumns.Contains(columnIndex);
    }

    internal IReadOnlySet<int> HiddenColumns => _hiddenColumns;

    /// <summary>
    /// Gets or sets whether this sheet is protected.
    /// When protected, worksheet's <c>&lt;sheetProtection&gt;</c> element is emitted on write.
    /// </summary>
    public bool SheetProtected
    {
        get => _sheetProtected;
        set => _sheetProtected = value;
    }

    public void protectSheet(bool protect)
    {
        _sheetProtected = protect;
    }

    public bool isSheetProtected() => _sheetProtected;

    // Auto filter
    private CellRangeAddress? _autoFilter;

    public void setAutoFilter(CellRangeAddress range)
    {
        ArgumentNullException.ThrowIfNull(range);
        _autoFilter = range;
    }

    public CellRangeAddress? getAutoFilter() => _autoFilter;

    internal CellRangeAddress? AutoFilter => _autoFilter;

    public IReadOnlyList<XSSFPivotTable> PivotTables => _pivotTables;

    /// <summary>
    /// Creates a pivot table on this sheet with data from the given source range.
    /// </summary>
    /// <param name="destCell">Destination cell for the top-left of the pivot table (e.g. "E5").</param>
    /// <param name="sourceRef">Source cell range (e.g. "A1:C100").</param>
    /// <param name="sourceSheetName">Name of the sheet containing the source data.</param>
    /// <returns>The created XSSFPivotTable.</returns>
    public XSSFPivotTable createPivotTable(string destCell, string sourceRef, string sourceSheetName)
    {
        ArgumentNullException.ThrowIfNull(destCell);
        ArgumentNullException.ThrowIfNull(sourceRef);
        ArgumentNullException.ThrowIfNull(sourceSheetName);

        int cacheId = _workbook.AllocatePivotCacheId();

        var pivotTable = new XSSFPivotTable
        {
            PivotTableIndex = _pivotTables.Count + 1,
            CacheId = cacheId,
            Name = "PivotTable" + (cacheId + 1).ToString(CultureInfo.InvariantCulture),
            DataCaption = "Values",
            DestinationCell = destCell,
            SourceAreaRef = sourceRef,
            SourceSheetName = sourceSheetName,
            Cache = new XSSFPivotCache(cacheId),
            CacheDefinition = new XSSFPivotCacheDefinition(cacheId, sourceSheetName, sourceRef),
            CacheRecords = new XSSFPivotCacheRecords(cacheId),
        };

        _pivotTables.Add(pivotTable);
        _workbook.RegisterPivotTable(pivotTable);
        return pivotTable;
    }

    IRow ISheet.createRow(int rownum) => createRow(rownum);

    IRow? ISheet.getRow(int rownum) => getRow(rownum);

    IWorkbook ISheet.getWorkbook() => getWorkbook();
}
