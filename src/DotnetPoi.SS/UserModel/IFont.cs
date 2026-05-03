namespace DotnetPoi.SS.UserModel;

public interface IFont
{
    int getIndex();
    bool getBold();
    void setBold(bool bold);
    short getColor();
    void setColor(short color);
    short getFontHeight();
    void setFontHeight(short height);
    short getFontHeightInPoints();
    void setFontHeightInPoints(short height);
    string getFontName();
    void setFontName(string? name);
    bool getItalic();
    void setItalic(bool italic);
    bool getStrikeout();
    void setStrikeout(bool strikeout);
    byte getUnderline();
    void setUnderline(byte underline);
}
