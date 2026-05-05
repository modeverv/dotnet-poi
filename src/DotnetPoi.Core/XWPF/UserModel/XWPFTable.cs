namespace DotnetPoi.XWPF.UserModel;

public sealed class XWPFTable
{
    private readonly List<XWPFTableRow> _rows = new();
    private readonly List<int> _gridColWidths = new();

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

    public void addGridCol(int width) => _gridColWidths.Add(width);
}

public sealed class XWPFTableRow
{
    private readonly List<XWPFTableCell> _cells = new();

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
}

public sealed class XWPFTableCell
{
    private readonly List<XWPFParagraph> _paragraphs = new();

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
}
