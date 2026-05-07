namespace DotnetPoi.XSLF.UserModel;

/// <summary>A row in a PPTX table.</summary>
public sealed class XSLFTableRow
{
    private readonly List<XSLFTableCell> _cells = new();

    internal XSLFTableRow() { }

    /// <summary>All cells in this row.</summary>
    public IReadOnlyList<XSLFTableCell> Cells => _cells;

    /// <summary>Creates and appends a new cell, returning it.</summary>
    public XSLFTableCell createCell()
    {
        var cell = new XSLFTableCell();
        _cells.Add(cell);
        return cell;
    }

    internal void AddCell(XSLFTableCell cell) => _cells.Add(cell);
}

/// <summary>A single cell in a PPTX table.</summary>
public sealed class XSLFTableCell
{
    private readonly List<XSLFTextParagraph> _paragraphs = new();

    internal XSLFTableCell() { }

    /// <summary>Text paragraphs in this cell.</summary>
    public IReadOnlyList<XSLFTextParagraph> Paragraphs => _paragraphs;

    /// <summary>Creates and appends a new paragraph, returning it.</summary>
    public XSLFTextParagraph addParagraph()
    {
        var p = new XSLFTextParagraph();
        _paragraphs.Add(p);
        return p;
    }

    internal void AddParagraph(XSLFTextParagraph p) => _paragraphs.Add(p);
}

/// <summary>Represents a table shape on a PPTX slide (p:graphicFrame > a:tbl).</summary>
public sealed class XSLFTable
{
    private readonly List<XSLFTableRow> _rows = new();
    private readonly List<long> _gridColWidths = new();

    internal XSLFTable(int shapeId)
    {
        ShapeId = shapeId;
    }

    /// <summary>Unique shape ID within the slide.</summary>
    public int ShapeId { get; }

    /// <summary>All rows in this table.</summary>
    public IReadOnlyList<XSLFTableRow> Rows => _rows;

    /// <summary>Grid column widths in EMU.</summary>
    public IReadOnlyList<long> GridColWidths => _gridColWidths;

    /// <summary>Adds a grid column width in EMU.</summary>
    public void addGridCol(long width) => _gridColWidths.Add(width);

    /// <summary>Creates and appends a new row, returning it.</summary>
    public XSLFTableRow createRow()
    {
        var row = new XSLFTableRow();
        _rows.Add(row);
        return row;
    }

    internal void AddRow(XSLFTableRow row) => _rows.Add(row);

    // --- anchor / position ---

    internal long AnchorX { get; private set; }
    internal long AnchorY { get; private set; }
    internal long AnchorCx { get; private set; }
    internal long AnchorCy { get; private set; }

    /// <summary>Sets position and size in EMU.</summary>
    public void setAnchor(long x, long y, long cx, long cy)
    {
        AnchorX = x;
        AnchorY = y;
        AnchorCx = cx;
        AnchorCy = cy;
    }

    public long getAnchorX() => AnchorX;
    public long getAnchorY() => AnchorY;
    public long getAnchorCx() => AnchorCx;
    public long getAnchorCy() => AnchorCy;
}
