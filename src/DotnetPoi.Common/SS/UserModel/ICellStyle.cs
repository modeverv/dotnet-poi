namespace DotnetPoi.SS.UserModel;

public interface ICellStyle
{
    int getIndex();
    short getDataFormat();
    string? getDataFormatString();
    void setDataFormat(short fmt);
    IFont getFont();
    void setFont(IFont? font);
    short getFillForegroundColor();
    void setFillForegroundColor(short fg);
    FillPatternType getFillPattern();
    void setFillPattern(FillPatternType pattern);
    BorderStyle getBorderTop();
    void setBorderTop(BorderStyle border);
    BorderStyle getBorderRight();
    void setBorderRight(BorderStyle border);
    BorderStyle getBorderBottom();
    void setBorderBottom(BorderStyle border);
    BorderStyle getBorderLeft();
    void setBorderLeft(BorderStyle border);
    HorizontalAlignment getAlignment();
    void setAlignment(HorizontalAlignment align);
    VerticalAlignment getVerticalAlignment();
    void setVerticalAlignment(VerticalAlignment align);
    bool getWrapText();
    void setWrapText(bool wrapped);
    short getIndention();
    void setIndention(short indent);
    short getRotation();
    void setRotation(short rotation);
}
