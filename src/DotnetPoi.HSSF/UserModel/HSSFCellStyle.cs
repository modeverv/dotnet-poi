using DotnetPoi.SS.UserModel;

namespace DotnetPoi.HSSF.UserModel;

public sealed class HSSFCellStyle : ICellStyle
{
    private readonly HSSFWorkbook _workbook;
    private readonly int _index;
    private short _dataFormat;
    private HSSFFont? _font;

    internal HSSFCellStyle(HSSFWorkbook workbook, int index)
    {
        _workbook = workbook;
        _index = index;
    }

    public int getIndex() => _index;

    public short getDataFormat() => _dataFormat;

    public string? getDataFormatString() => _workbook.createDataFormat().getFormat(_dataFormat);

    public void setDataFormat(short fmt) => _dataFormat = fmt;

    public HSSFFont getFont() => _font ?? _workbook.getFontAt(0);

    public void setFont(IFont? font) => _font = font as HSSFFont;

    IFont ICellStyle.getFont() => getFont();

    public short getFillForegroundColor() => 0;

    public void setFillForegroundColor(short fg)
    {
        // TODO: [dotnet-poi] Not yet ported
        // Original: poi/poi/src/main/java/org/apache/poi/hssf/usermodel/HSSFCellStyle.java
        // Reason: Phase 6 xls bootstrap only persists basic cell values.
        // Issue: Phase 6 HSSF styles backlog
    }

    public FillPatternType getFillPattern() => FillPatternType.NoFill;

    public void setFillPattern(FillPatternType pattern)
    {
    }

    public BorderStyle getBorderTop() => BorderStyle.None;

    public void setBorderTop(BorderStyle border)
    {
    }

    public BorderStyle getBorderRight() => BorderStyle.None;

    public void setBorderRight(BorderStyle border)
    {
    }

    public BorderStyle getBorderBottom() => BorderStyle.None;

    public void setBorderBottom(BorderStyle border)
    {
    }

    public BorderStyle getBorderLeft() => BorderStyle.None;

    public void setBorderLeft(BorderStyle border)
    {
    }
}
