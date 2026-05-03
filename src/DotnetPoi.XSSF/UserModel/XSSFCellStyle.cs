using DotnetPoi.SS.UserModel;

namespace DotnetPoi.XSSF.UserModel;

public sealed class XSSFCellStyle : ICellStyle
{
    private readonly XSSFWorkbook _workbook;

    internal XSSFCellStyle(XSSFWorkbook workbook, int index)
    {
        _workbook = workbook;
        Index = index;
    }

    internal int Index { get; }

    internal int FontId { get; private set; }

    internal int NumFmtId { get; private set; }

    internal int FillId { get; private set; }

    internal int BorderId { get; private set; }

    internal bool ApplyFont { get; private set; }

    internal bool ApplyNumberFormat { get; private set; }

    internal bool ApplyFill { get; private set; }

    internal bool ApplyBorder { get; private set; }

    internal short? FillForegroundColor { get; private set; }

    internal FillPatternType FillPattern { get; private set; } = FillPatternType.NoFill;

    internal BorderStyle BorderTop { get; private set; }

    internal BorderStyle BorderRight { get; private set; }

    internal BorderStyle BorderBottom { get; private set; }

    internal BorderStyle BorderLeft { get; private set; }

    internal XSSFWorkbook Workbook => _workbook;

    public int getIndex()
    {
        return Index;
    }

    public short getDataFormat()
    {
        return (short)NumFmtId;
    }

    public string? getDataFormatString()
    {
        return _workbook.createDataFormat().getFormat((short)NumFmtId);
    }

    public void setDataFormat(short fmt)
    {
        setDataFormat(fmt & 0xffff);
    }

    public void setDataFormat(int fmt)
    {
        NumFmtId = fmt;
        ApplyNumberFormat = true;
    }

    public XSSFFont getFont()
    {
        return _workbook.getFontAt(FontId);
    }

    public void setFont(XSSFFont? font)
    {
        if (font is null)
        {
            FontId = 0;
            ApplyFont = false;
            return;
        }

        _workbook.VerifyFontBelongsToWorkbook(font);
        FontId = font.getIndex();
        ApplyFont = true;
    }

    public void setFont(IFont? font)
    {
        if (font is not null && font is not XSSFFont)
        {
            throw new ArgumentException("Font must be an XSSFFont created by this workbook.", nameof(font));
        }

        setFont(font as XSSFFont);
    }

    public void setFillForegroundColor(short fg)
    {
        FillForegroundColor = fg;
        RegisterFill();
    }

    public short getFillForegroundColor()
    {
        return FillForegroundColor ?? 0;
    }

    public void setFillPattern(FillPatternType pattern)
    {
        FillPattern = pattern;
        RegisterFill();
    }

    public FillPatternType getFillPattern()
    {
        return FillPattern;
    }

    public void setBorderTop(BorderStyle border)
    {
        BorderTop = border;
        RegisterBorder();
    }

    public BorderStyle getBorderTop()
    {
        return BorderTop;
    }

    public void setBorderRight(BorderStyle border)
    {
        BorderRight = border;
        RegisterBorder();
    }

    public BorderStyle getBorderRight()
    {
        return BorderRight;
    }

    public void setBorderBottom(BorderStyle border)
    {
        BorderBottom = border;
        RegisterBorder();
    }

    public BorderStyle getBorderBottom()
    {
        return BorderBottom;
    }

    public void setBorderLeft(BorderStyle border)
    {
        BorderLeft = border;
        RegisterBorder();
    }

    public BorderStyle getBorderLeft()
    {
        return BorderLeft;
    }

    IFont ICellStyle.getFont() => getFont();

    private void RegisterFill()
    {
        if (FillId == 0)
        {
            FillId = _workbook.GetOrAddFill(this);
        }
        ApplyFill = true;
    }

    private void RegisterBorder()
    {
        if (BorderId == 0)
        {
            BorderId = _workbook.GetOrAddBorder(this);
        }
        ApplyBorder = true;
    }
}
