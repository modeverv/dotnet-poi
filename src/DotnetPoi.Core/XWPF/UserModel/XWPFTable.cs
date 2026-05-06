namespace DotnetPoi.XWPF.UserModel;

public sealed class XWPFTable
{
    private readonly List<XWPFTableRow> _rows = new();
    private readonly List<int> _gridColWidths = new();

    // tblPr modeled properties
    private long _tableWidth;
    private string? _tableWidthType; // "dxa", "pct", "auto"
    private string? _tableStyle;

    // Raw XML preservation for unmodeled tblPr children (borders, shading, cellMar, layout, look, etc.)
    private readonly List<string> _preservedRawTblPrChildren = new();

    internal XWPFTable(XWPFDocument document)
    {
        Document = document;
    }

    internal XWPFDocument Document { get; }

    public XWPFTableRow createRow()
    {
        var row = new XWPFTableRow(this);
        _rows.Add(row);
        return row;
    }

    public IReadOnlyList<XWPFTableRow> getRows() => _rows;
    internal IReadOnlyList<XWPFTableRow> Rows => _rows;
    internal IReadOnlyList<int> GridColWidths => _gridColWidths;
    internal IReadOnlyList<string> PreservedRawTblPrChildren => _preservedRawTblPrChildren;

    public void addGridCol(int width) => _gridColWidths.Add(width);

    public long getWidth() => _tableWidth;
    public string? getWidthType() => _tableWidthType;
    public void setWidth(long width, string type) { _tableWidth = width; _tableWidthType = type; }

    internal void setTableStyle(string? style) => _tableStyle = style;
    internal string? getTableStyle() => _tableStyle;
    internal void addPreservedRawTblPrChild(string raw) => _preservedRawTblPrChildren.Add(raw);
}

public sealed class XWPFTableRow
{
    private readonly List<XWPFTableCell> _cells = new();

    // trPr modeled properties
    private long _height;
    private string? _heightRule; // "atLeast", "exact", "auto"
    private bool _isHeader;

    // Raw XML preservation for unmodeled trPr children
    private readonly List<string> _preservedRawTrPrChildren = new();

    internal XWPFTableRow(XWPFTable table)
    {
        Table = table;
    }

    internal XWPFTable Table { get; }

    public XWPFTableCell createCell()
    {
        var cell = new XWPFTableCell(this);
        _cells.Add(cell);
        return cell;
    }

    public IReadOnlyList<XWPFTableCell> getCells() => _cells;
    internal IReadOnlyList<XWPFTableCell> Cells => _cells;

    public long getHeight() => _height;
    public string? getHeightRule() => _heightRule;
    public void setHeight(long height, string rule) { _height = height; _heightRule = rule; }

    public bool isHeader() => _isHeader;
    public void setHeader(bool header) => _isHeader = header;

    internal IReadOnlyList<string> PreservedRawTrPrChildren => _preservedRawTrPrChildren;
    internal void addPreservedRawTrPrChild(string raw) => _preservedRawTrPrChildren.Add(raw);
}

public sealed class XWPFTableCell
{
    private readonly List<XWPFParagraph> _paragraphs = new();

    // tcPr modeled properties
    private long _cellWidth;
    private string? _cellWidthType; // "dxa", "pct", "auto"
    private int _gridSpan = 1;
    private string? _vMerge; // "restart", "continue", or null
    private string? _hMerge; // "restart", "continue", or null
    private string? _vAlign; // "top", "center", "bottom"

    // Raw XML preservation for unmodeled tcPr children (borders, shading, cellMar, textDirection, etc.)
    private readonly List<string> _preservedRawTcPrChildren = new();

    internal XWPFTableCell(XWPFTableRow row)
    {
        Row = row;
    }

    internal XWPFTableRow Row { get; }
    internal IReadOnlyList<XWPFParagraph> Paragraphs => _paragraphs;

    public XWPFParagraph addParagraph()
    {
        var para = new XWPFParagraph(Row.Table.Document);
        _paragraphs.Add(para);
        return para;
    }

    public IReadOnlyList<XWPFParagraph> getParagraphs() => _paragraphs;

    // Cell width
    public long getWidth() => _cellWidth;
    public string? getWidthType() => _cellWidthType;
    public void setWidth(long width, string type) { _cellWidth = width; _cellWidthType = type; }

    // Grid span (horizontal merge — number of columns spanned)
    public int getGridSpan() => _gridSpan;
    public void setGridSpan(int span) => _gridSpan = span;

    // Vertical merge
    public string? getVMerge() => _vMerge;
    public void setVMerge(string? vMerge) => _vMerge = vMerge;

    // Horizontal merge (legacy; gridSpan is preferred)
    public string? getHMerge() => _hMerge;
    public void setHMerge(string? hMerge) => _hMerge = hMerge;

    // Vertical alignment
    public string? getVAlign() => _vAlign;
    public void setVAlign(string? vAlign) => _vAlign = vAlign;

    internal IReadOnlyList<string> PreservedRawTcPrChildren => _preservedRawTcPrChildren;
    internal void addPreservedRawTcPrChild(string raw) => _preservedRawTcPrChildren.Add(raw);
}
