namespace DotnetPoi.XSSF.UserModel;

/// <summary>
/// Cell comment for XSSF worksheets.
/// Ported from org.apache.poi.xssf.usermodel.XSSFComment at the public API level.
/// </summary>
public sealed class XSSFComment
{
    private readonly XSSFSheet _sheet;
    private XSSFRichTextString? _text;
    private string _author;
    private bool _visible;
    private XSSFClientAnchor? _anchor;

    internal XSSFComment(XSSFSheet sheet, int row, int column, string author, XSSFRichTextString? text, XSSFClientAnchor? anchor, bool visible)
    {
        _sheet = sheet;
        Row = row;
        Column = column;
        _author = author;
        _text = text;
        _anchor = anchor;
        _visible = visible;
    }

    public int Row { get; private set; }

    public int Column { get; private set; }

    public string getAuthor() => _author;

    public void setAuthor(string author)
    {
        _author = author ?? string.Empty;
        _sheet.RegisterComment(this);
    }

    public int getRow() => Row;

    public void setRow(int row) => setAddress(row, Column);

    public int getColumn() => Column;

    public void setColumn(int col) => setAddress(Row, col);

    public string getAddress() => CellRef;

    public void setAddress(int row, int col)
    {
        var oldRow = Row;
        var oldColumn = Column;
        Row = row;
        Column = col;

        if (_anchor is not null)
        {
            var rowSpan = Math.Max(1, _anchor.Row2 - _anchor.Row1);
            var colSpan = Math.Max(1, _anchor.Col2 - _anchor.Col1);
            _anchor = new XSSFClientAnchor(_anchor.Dx1, _anchor.Dy1, _anchor.Dx2, _anchor.Dy2, col, row, col + colSpan, row + rowSpan);
        }

        _sheet.MoveComment(oldRow, oldColumn, this);
    }

    public XSSFRichTextString? getString() => _text;

    public void setString(string text) => setString(new XSSFRichTextString(text));

    public void setString(XSSFRichTextString? text)
    {
        _text = text;
        _sheet.RegisterComment(this);
    }

    public bool isVisible() => _visible;

    public void setVisible(bool visible)
    {
        _visible = visible;
        _sheet.RegisterComment(this);
    }

    public XSSFClientAnchor? getClientAnchor() => _anchor;

    internal void SetClientAnchor(XSSFClientAnchor? anchor) => _anchor = anchor;

    internal string CellRef => XSSFHyperlink.FormatCellRef(Row, Column);
}
