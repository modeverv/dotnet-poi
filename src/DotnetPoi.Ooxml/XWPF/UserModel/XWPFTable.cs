namespace DotnetPoi.XWPF.UserModel;

public sealed class XWPFTable
{
    public enum XWPFBorderType
    {
        Nil,
        None,
        Single,
        Thick,
        Double,
        Dotted,
        Dashed,
        DotDash,
        DotDotDash,
        Triple,
        ThinThickSmallGap,
        ThickThinSmallGap,
        ThinThickThinSmallGap,
        ThinThickMediumGap,
        ThickThinMediumGap,
        ThinThickThinMediumGap,
        ThinThickLargeGap,
        ThickThinLargeGap,
        ThinThickThinLargeGap,
        Wave,
        DoubleWave,
        DashSmallGap,
        DashDotStroked,
        ThreeDEmboss,
        ThreeDEngrave,
        Outset,
        Inset
    }

    public sealed record TableBorder(XWPFBorderType Type, int Size, int Space, string Color);

    private readonly List<XWPFTableRow> _rows = new();
    private readonly List<int> _gridColWidths = new();
    private readonly Dictionary<TableBorderPosition, TableBorder> _borders = new();

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
    internal IReadOnlyDictionary<TableBorderPosition, TableBorder> Borders => _borders;

    public void addGridCol(int width) => _gridColWidths.Add(width);

    public long getWidth() => _tableWidth;
    public string? getWidthType() => _tableWidthType;
    public void setWidth(long width, string type) { _tableWidth = width; _tableWidthType = type; }

    internal void setTableStyle(string? style) => _tableStyle = style;
    internal string? getTableStyle() => _tableStyle;
    internal void addPreservedRawTblPrChild(string raw) => _preservedRawTblPrChildren.Add(raw);

    public void mergeCellsHorizontally(int row, int fromCell, int toCell)
    {
        if (row < 0 || row >= _rows.Count)
            throw new ArgumentOutOfRangeException(nameof(row));
        if (fromCell < 0 || toCell < fromCell)
            throw new ArgumentOutOfRangeException(nameof(fromCell));

        var cells = _rows[row].Cells;
        if (toCell >= cells.Count)
            throw new ArgumentOutOfRangeException(nameof(toCell));

        cells[fromCell].setGridSpan(toCell - fromCell + 1);
        for (var i = fromCell + 1; i <= toCell; i++)
        {
            cells[i].setHMerge("continue");
        }
    }

    public void mergeCellsVertically(int column, int fromRow, int toRow)
    {
        if (column < 0)
            throw new ArgumentOutOfRangeException(nameof(column));
        if (fromRow < 0 || toRow < fromRow)
            throw new ArgumentOutOfRangeException(nameof(fromRow));
        if (toRow >= _rows.Count)
            throw new ArgumentOutOfRangeException(nameof(toRow));

        for (var r = fromRow; r <= toRow; r++)
        {
            if (column >= _rows[r].Cells.Count)
                throw new ArgumentOutOfRangeException(nameof(column));
            _rows[r].Cells[column].setVMerge(r == fromRow ? "restart" : "continue");
        }
    }

    public void setTopBorder(XWPFBorderType type, int size, int space, string rgbColor) =>
        SetBorder(TableBorderPosition.Top, type, size, space, rgbColor);

    public void setBottomBorder(XWPFBorderType type, int size, int space, string rgbColor) =>
        SetBorder(TableBorderPosition.Bottom, type, size, space, rgbColor);

    public void setLeftBorder(XWPFBorderType type, int size, int space, string rgbColor) =>
        SetBorder(TableBorderPosition.Left, type, size, space, rgbColor);

    public void setRightBorder(XWPFBorderType type, int size, int space, string rgbColor) =>
        SetBorder(TableBorderPosition.Right, type, size, space, rgbColor);

    public void setInsideHBorder(XWPFBorderType type, int size, int space, string rgbColor) =>
        SetBorder(TableBorderPosition.InsideH, type, size, space, rgbColor);

    public void setInsideVBorder(XWPFBorderType type, int size, int space, string rgbColor) =>
        SetBorder(TableBorderPosition.InsideV, type, size, space, rgbColor);

    public XWPFBorderType? getTopBorderType() => GetBorderType(TableBorderPosition.Top);
    public XWPFBorderType? getBottomBorderType() => GetBorderType(TableBorderPosition.Bottom);
    public XWPFBorderType? getLeftBorderType() => GetBorderType(TableBorderPosition.Left);
    public XWPFBorderType? getRightBorderType() => GetBorderType(TableBorderPosition.Right);
    public XWPFBorderType? getInsideHBorderType() => GetBorderType(TableBorderPosition.InsideH);
    public XWPFBorderType? getInsideVBorderType() => GetBorderType(TableBorderPosition.InsideV);

    public int getTopBorderSize() => GetBorderSize(TableBorderPosition.Top);
    public int getBottomBorderSize() => GetBorderSize(TableBorderPosition.Bottom);
    public int getLeftBorderSize() => GetBorderSize(TableBorderPosition.Left);
    public int getRightBorderSize() => GetBorderSize(TableBorderPosition.Right);
    public int getInsideHBorderSize() => GetBorderSize(TableBorderPosition.InsideH);
    public int getInsideVBorderSize() => GetBorderSize(TableBorderPosition.InsideV);

    public int getTopBorderSpace() => GetBorderSpace(TableBorderPosition.Top);
    public int getBottomBorderSpace() => GetBorderSpace(TableBorderPosition.Bottom);
    public int getLeftBorderSpace() => GetBorderSpace(TableBorderPosition.Left);
    public int getRightBorderSpace() => GetBorderSpace(TableBorderPosition.Right);
    public int getInsideHBorderSpace() => GetBorderSpace(TableBorderPosition.InsideH);
    public int getInsideVBorderSpace() => GetBorderSpace(TableBorderPosition.InsideV);

    public string? getTopBorderColor() => GetBorderColor(TableBorderPosition.Top);
    public string? getBottomBorderColor() => GetBorderColor(TableBorderPosition.Bottom);
    public string? getLeftBorderColor() => GetBorderColor(TableBorderPosition.Left);
    public string? getRightBorderColor() => GetBorderColor(TableBorderPosition.Right);
    public string? getInsideHBorderColor() => GetBorderColor(TableBorderPosition.InsideH);
    public string? getInsideVBorderColor() => GetBorderColor(TableBorderPosition.InsideV);

    public void removeBorders() => _borders.Clear();
    public void removeTopBorder() => _borders.Remove(TableBorderPosition.Top);
    public void removeBottomBorder() => _borders.Remove(TableBorderPosition.Bottom);
    public void removeLeftBorder() => _borders.Remove(TableBorderPosition.Left);
    public void removeRightBorder() => _borders.Remove(TableBorderPosition.Right);
    public void removeInsideHBorder() => _borders.Remove(TableBorderPosition.InsideH);
    public void removeInsideVBorder() => _borders.Remove(TableBorderPosition.InsideV);

    internal void SetBorder(TableBorderPosition position, XWPFBorderType type, int size, int space, string? rgbColor)
    {
        _borders[position] = new TableBorder(type, size, space, rgbColor ?? "auto");
    }

    private XWPFBorderType? GetBorderType(TableBorderPosition position) =>
        _borders.TryGetValue(position, out var border) ? border.Type : null;

    private int GetBorderSize(TableBorderPosition position) =>
        _borders.TryGetValue(position, out var border) ? border.Size : -1;

    private int GetBorderSpace(TableBorderPosition position) =>
        _borders.TryGetValue(position, out var border) ? border.Space : -1;

    private string? GetBorderColor(TableBorderPosition position) =>
        _borders.TryGetValue(position, out var border) ? border.Color : null;
}

internal enum TableBorderPosition
{
    Top,
    Bottom,
    Left,
    Right,
    InsideH,
    InsideV
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
    public enum XWPFVertAlign
    {
        Top,
        Center,
        Both,
        Bottom
    }

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
    public void setWidth(string widthValue)
    {
        if (string.Equals(widthValue, "auto", StringComparison.OrdinalIgnoreCase))
        {
            setWidth(0, "auto");
            return;
        }

        if (widthValue.EndsWith("%", StringComparison.Ordinal))
        {
            var number = widthValue.Substring(0, widthValue.Length - 1);
            if (!double.TryParse(number, System.Globalization.NumberStyles.Float,
                    System.Globalization.CultureInfo.InvariantCulture, out var percentage))
                throw new ArgumentException("Width percentage must be numeric.", nameof(widthValue));
            setWidth((long)Math.Round(percentage * 50.0), "pct");
            return;
        }

        if (!long.TryParse(widthValue, System.Globalization.NumberStyles.Integer,
                System.Globalization.CultureInfo.InvariantCulture, out var width))
            throw new ArgumentException("Width must be \"auto\", a twips integer, or a percentage like \"33.3%\".", nameof(widthValue));
        setWidth(width, "dxa");
    }

    // Grid span (horizontal merge — number of columns spanned)
    public int getGridSpan() => _gridSpan;
    public void setGridSpan(int span) => _gridSpan = span <= 1 ? 1 : span;

    // Vertical merge
    public string? getVMerge() => _vMerge;
    public void setVMerge(string? vMerge) => _vMerge = vMerge;

    // Horizontal merge (legacy; gridSpan is preferred)
    public string? getHMerge() => _hMerge;
    public void setHMerge(string? hMerge) => _hMerge = hMerge;

    // Vertical alignment
    public string? getVAlign() => _vAlign;
    public void setVAlign(string? vAlign) => _vAlign = vAlign;
    public XWPFVertAlign? getVerticalAlignment()
    {
        return _vAlign switch
        {
            "top" => XWPFVertAlign.Top,
            "center" => XWPFVertAlign.Center,
            "both" => XWPFVertAlign.Both,
            "bottom" => XWPFVertAlign.Bottom,
            _ => null
        };
    }

    public void setVerticalAlignment(XWPFVertAlign vAlign)
    {
        _vAlign = vAlign switch
        {
            XWPFVertAlign.Top => "top",
            XWPFVertAlign.Center => "center",
            XWPFVertAlign.Both => "both",
            XWPFVertAlign.Bottom => "bottom",
            _ => null
        };
    }

    internal IReadOnlyList<string> PreservedRawTcPrChildren => _preservedRawTcPrChildren;
    internal void addPreservedRawTcPrChild(string raw) => _preservedRawTcPrChildren.Add(raw);
}
