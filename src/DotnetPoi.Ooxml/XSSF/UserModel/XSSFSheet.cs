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
    private readonly SortedDictionary<(int Row, int Column), XSSFComment> _comments = new();
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

        // POI semantics: return existing row if already created
        if (_rows.TryGetValue(rownum, out var existing))
            return existing;

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

    internal XSSFDrawing? Drawing =>
        _drawing is not null && (_drawing.Pictures.Count > 0 || _drawing.PreservedRawAnchors.Count > 0)
            ? _drawing
            : null;

    internal bool HasComments => _comments.Count > 0;

    internal int CommentsIndex => SheetIndex;

    internal IReadOnlyCollection<XSSFComment> Comments => _comments.Values;

    public XSSFComment? getCellComment(int row, int column)
    {
        return _comments.TryGetValue((row, column), out var comment) ? comment : null;
    }

    public XSSFComment? findCellComment(int row, int column) => getCellComment(row, column);

    public IReadOnlyDictionary<string, XSSFComment> getCellComments()
    {
        return _comments.ToDictionary(
            pair => XSSFHyperlink.FormatCellRef(pair.Key.Row, pair.Key.Column),
            pair => pair.Value);
    }

    internal void RegisterComment(XSSFComment comment)
    {
        Guard.ThrowIfNull(comment, nameof(comment));
        _comments[(comment.getRow(), comment.getColumn())] = comment;
    }

    internal void MoveComment(int oldRow, int oldColumn, XSSFComment comment)
    {
        _comments.Remove((oldRow, oldColumn));
        RegisterComment(comment);
    }

    internal XSSFComment CreateCellComment(XSSFClientAnchor anchor)
    {
        Guard.ThrowIfNull(anchor, nameof(anchor));
        if (_comments.ContainsKey((anchor.Row1, anchor.Col1)))
            throw new ArgumentException("Multiple cell comments in one cell are not allowed.");

        var comment = new XSSFComment(this, anchor.Row1, anchor.Col1, string.Empty, null, anchor, visible: false);
        RegisterComment(comment);
        return comment;
    }

    internal void SetCellComment(XSSFCell cell, XSSFComment? comment)
    {
        var key = (cell.getRowIndex(), cell.getColumnIndex());
        if (comment is null)
        {
            _comments.Remove(key);
            return;
        }

        comment.setAddress(cell.getRowIndex(), cell.getColumnIndex());
        RegisterComment(comment);
    }

    public void addMergedRegion(CellRangeAddress region)
    {
        Guard.ThrowIfNull(region, nameof(region));
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
        Guard.ThrowIfNull(link, nameof(link));
        _hyperlinks.Add(link);
    }

    public void AddDataValidation(XSSFDataValidation validation)
    {
        Guard.ThrowIfNull(validation, nameof(validation));
        _dataValidations.Add(validation);
    }

    public IReadOnlyList<XSSFConditionalFormatting> ConditionalFormatting => _condFormatting;

    public void AddConditionalFormatting(XSSFConditionalFormatting cf)
    {
        Guard.ThrowIfNull(cf, nameof(cf));
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

    internal IReadOnlyCollection<int> HiddenColumns => _hiddenColumns;

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
        Guard.ThrowIfNull(range, nameof(range));
        _autoFilter = range;
    }

    public CellRangeAddress? getAutoFilter() => _autoFilter;

    internal CellRangeAddress? AutoFilter => _autoFilter;

    // Active cell / selection
    private string? _activeCell;
    private bool _selected;

    public void setActiveCell(string cellRef)
    {
        _activeCell = cellRef ?? throw new ArgumentNullException(nameof(cellRef));
    }

    public string? getActiveCell() => _activeCell;

    public void setSelected(bool selected)
    {
        _selected = selected;
    }

    public bool isSelected() => _selected;

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
        Guard.ThrowIfNull(destCell, nameof(destCell));
        Guard.ThrowIfNull(sourceRef, nameof(sourceRef));
        Guard.ThrowIfNull(sourceSheetName, nameof(sourceSheetName));

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

    void ISheet.setActiveCell(string cell) => setActiveCell(cell);

    string? ISheet.getActiveCell() => getActiveCell();

    void ISheet.setSelected(bool selected) => setSelected(selected);

    bool ISheet.isSelected() => isSelected();
}
