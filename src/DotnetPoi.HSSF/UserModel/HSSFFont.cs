using DotnetPoi.SS.UserModel;

namespace DotnetPoi.HSSF.UserModel;

public sealed class HSSFFont : IFont
{
    private readonly int _index;
    private bool _bold;
    private short _color;
    private short _height = 200;
    private string _fontName = "Arial";
    private bool _italic;
    private bool _strikeout;
    private byte _underline;

    internal HSSFFont(int index)
    {
        _index = index;
    }

    public int getIndex() => _index;

    public bool getBold() => _bold;

    public void setBold(bool bold) => _bold = bold;

    public short getColor() => _color;

    public void setColor(short color) => _color = color;

    public short getFontHeight() => _height;

    public void setFontHeight(short height) => _height = height;

    public short getFontHeightInPoints() => (short)(_height / 20);

    public void setFontHeightInPoints(short height) => _height = (short)(height * 20);

    public string getFontName() => _fontName;

    public void setFontName(string? name) => _fontName = name ?? string.Empty;

    public bool getItalic() => _italic;

    public void setItalic(bool italic) => _italic = italic;

    public bool getStrikeout() => _strikeout;

    public void setStrikeout(bool strikeout) => _strikeout = strikeout;

    public byte getUnderline() => _underline;

    public void setUnderline(byte underline) => _underline = underline;
}
