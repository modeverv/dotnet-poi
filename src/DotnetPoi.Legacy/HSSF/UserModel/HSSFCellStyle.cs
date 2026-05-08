using DotnetPoi.SS.UserModel;

namespace DotnetPoi.HSSF.UserModel;

public sealed class HSSFCellStyle : ICellStyle
{
    private readonly HSSFWorkbook _workbook;
    private readonly int _index;
    private short _parentIndex = 0x0FFF;
    private short _dataFormat;
    private HSSFFont? _font;
    private HorizontalAlignment _alignment = HorizontalAlignment.General;
    private VerticalAlignment _verticalAlignment = VerticalAlignment.Bottom;
    private bool _wrapText;
    private short _indention;
    private short _rotation;
    private BorderStyle _borderTop = BorderStyle.None;
    private BorderStyle _borderRight = BorderStyle.None;
    private BorderStyle _borderBottom = BorderStyle.None;
    private BorderStyle _borderLeft = BorderStyle.None;
    private short _fillForegroundColor;
    private FillPatternType _fillPattern = FillPatternType.NoFill;

    internal HSSFCellStyle(HSSFWorkbook workbook, int index)
    {
        _workbook = workbook;
        _index = index;
    }

    public int getIndex() => _index;

    public short getParentIndex() => _parentIndex;

    internal void setParentIndex(short parentIndex) => _parentIndex = parentIndex;

    public short getDataFormat() => _dataFormat;

    public string? getDataFormatString() => _workbook.createDataFormat().getFormat(_dataFormat);

    public void setDataFormat(short fmt) => _dataFormat = fmt;

    public HSSFFont getFont() => _font ?? _workbook.getFontAt(0);

    public void setFont(IFont? font) => _font = font as HSSFFont;

    IFont ICellStyle.getFont() => getFont();

    public short getFillForegroundColor() => _fillForegroundColor;

    public void setFillForegroundColor(short fg) => _fillForegroundColor = fg;

    public FillPatternType getFillPattern() => _fillPattern;

    public void setFillPattern(FillPatternType pattern) => _fillPattern = pattern;

    public BorderStyle getBorderTop() => _borderTop;

    public void setBorderTop(BorderStyle border) => _borderTop = border;

    public BorderStyle getBorderRight() => _borderRight;

    public void setBorderRight(BorderStyle border) => _borderRight = border;

    public BorderStyle getBorderBottom() => _borderBottom;

    public void setBorderBottom(BorderStyle border) => _borderBottom = border;

    public BorderStyle getBorderLeft() => _borderLeft;

    public void setBorderLeft(BorderStyle border) => _borderLeft = border;

    public HorizontalAlignment getAlignment() => _alignment;

    public void setAlignment(HorizontalAlignment align) => _alignment = align;

    public VerticalAlignment getVerticalAlignment() => _verticalAlignment;

    public void setVerticalAlignment(VerticalAlignment align) => _verticalAlignment = align;

    public bool getWrapText() => _wrapText;

    public void setWrapText(bool wrapped) => _wrapText = wrapped;

    public short getIndention() => _indention;

    public void setIndention(short indent) => _indention = indent;

    public short getRotation() => _rotation;

    public void setRotation(short rotation) => _rotation = rotation;

    internal int FontBiffIndex => _font?.getIndex() ?? 0;
}
