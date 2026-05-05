using DotnetPoi.SS.UserModel;

namespace DotnetPoi.XSSF.UserModel;

public sealed class XSSFCellStyle : ICellStyle
{
    private readonly XSSFWorkbook _workbook;
    private bool _fillRegistered;
    private bool _borderRegistered;

    internal XSSFCellStyle(XSSFWorkbook workbook, int index)
    {
        _workbook = workbook;
        Index = index;
    }

    internal int Index { get; }

    internal int FontId { get; set; }

    internal int NumFmtId { get; set; }

    internal int FillId { get; set; }

    internal int BorderId { get; set; }

    internal bool ApplyFont { get; set; }

    internal bool ApplyNumberFormat { get; set; }

    internal bool ApplyFill { get; set; }

    internal bool ApplyBorder { get; set; }

    internal bool ApplyAlignment { get; set; }

    internal short? FillForegroundColor { get; set; }

    internal FillPatternType FillPattern { get; set; } = FillPatternType.NoFill;

    internal BorderStyle BorderTop { get; set; }

    internal BorderStyle BorderRight { get; set; }

    internal BorderStyle BorderBottom { get; set; }

    internal BorderStyle BorderLeft { get; set; }

    internal HorizontalAlignment AlignmentValue { get; set; } = HorizontalAlignment.General;

    internal VerticalAlignment VerticalAlignmentValue { get; set; } = VerticalAlignment.Bottom;

    internal bool WrapTextEnabled { get; set; }

    internal short IndentLevel { get; set; }

    internal short TextRotation { get; set; }

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
        ApplyFill = true;
        if (!_fillRegistered) { FillId = _workbook.GetOrAddFill(this); _fillRegistered = true; }
    }

    public short getFillForegroundColor()
    {
        return FillForegroundColor ?? 0;
    }

    public void setFillPattern(FillPatternType pattern)
    {
        FillPattern = pattern;
        ApplyFill = true;
        if (!_fillRegistered) { FillId = _workbook.GetOrAddFill(this); _fillRegistered = true; }
    }

    public FillPatternType getFillPattern()
    {
        return FillPattern;
    }

    public void setBorderTop(BorderStyle border)
    {
        BorderTop = border;
        ApplyBorder = true;
        if (!_borderRegistered) { BorderId = _workbook.GetOrAddBorder(this); _borderRegistered = true; }
    }

    public BorderStyle getBorderTop()
    {
        return ApplyBorder ? BorderTop : BorderStyle.None;
    }

    public void setBorderRight(BorderStyle border)
    {
        BorderRight = border;
        ApplyBorder = true;
        if (!_borderRegistered) { BorderId = _workbook.GetOrAddBorder(this); _borderRegistered = true; }
    }

    public BorderStyle getBorderRight()
    {
        return ApplyBorder ? BorderRight : BorderStyle.None;
    }

    public void setBorderBottom(BorderStyle border)
    {
        BorderBottom = border;
        ApplyBorder = true;
        if (!_borderRegistered) { BorderId = _workbook.GetOrAddBorder(this); _borderRegistered = true; }
    }

    public BorderStyle getBorderBottom()
    {
        return ApplyBorder ? BorderBottom : BorderStyle.None;
    }

    public void setBorderLeft(BorderStyle border)
    {
        BorderLeft = border;
        ApplyBorder = true;
        if (!_borderRegistered) { BorderId = _workbook.GetOrAddBorder(this); _borderRegistered = true; }
    }

    public BorderStyle getBorderLeft()
    {
        return ApplyBorder ? BorderLeft : BorderStyle.None;
    }

    public HorizontalAlignment getAlignment()
    {
        return ApplyAlignment ? AlignmentValue : HorizontalAlignment.General;
    }

    public void setAlignment(HorizontalAlignment align)
    {
        AlignmentValue = align;
        ApplyAlignment = true;
    }

    public VerticalAlignment getVerticalAlignment()
    {
        return ApplyAlignment ? VerticalAlignmentValue : VerticalAlignment.Bottom;
    }

    public void setVerticalAlignment(VerticalAlignment align)
    {
        VerticalAlignmentValue = align;
        ApplyAlignment = true;
    }

    public bool getWrapText()
    {
        return ApplyAlignment && WrapTextEnabled;
    }

    public void setWrapText(bool wrapped)
    {
        WrapTextEnabled = wrapped;
        ApplyAlignment = true;
    }

    public short getIndention()
    {
        return IndentLevel;
    }

    public void setIndention(short indent)
    {
        IndentLevel = indent;
        ApplyAlignment = true;
    }

    public short getRotation()
    {
        return TextRotation;
    }

    public void setRotation(short rotation)
    {
        TextRotation = rotation;
        ApplyAlignment = true;
    }

    IFont ICellStyle.getFont() => getFont();
}
