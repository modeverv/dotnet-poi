using System.Globalization;
using DotnetPoi.SS.UserModel;

namespace DotnetPoi.XSSF.UserModel;

public sealed class XSSFFont : IFont
{
    public const string DEFAULT_FONT_NAME = "Calibri";
    public const short DEFAULT_FONT_SIZE = 11;
    public const short DEFAULT_FONT_COLOR = (short)IndexedColors.Black;
    public const short COLOR_NORMAL = short.MaxValue;

    internal XSSFFont(int index)
    {
        Index = index;
    }

    internal int Index { get; set; }

    internal bool Bold { get; private set; }

    internal bool Italic { get; private set; }

    internal bool Strikeout { get; private set; }

    internal byte Underline { get; private set; }

    internal short Color { get; private set; } = DEFAULT_FONT_COLOR;

    internal double FontHeightInPoints { get; private set; } = DEFAULT_FONT_SIZE;

    internal string FontName { get; private set; } = DEFAULT_FONT_NAME;

    public int getIndex()
    {
        return Index;
    }

    public bool getBold()
    {
        return Bold;
    }

    public void setBold(bool bold)
    {
        Bold = bold;
    }

    public short getColor()
    {
        return Color;
    }

    public void setColor(short color)
    {
        Color = color == COLOR_NORMAL ? DEFAULT_FONT_COLOR : color;
    }

    public short getFontHeight()
    {
        return (short)(FontHeightInPoints * 20);
    }

    public short getFontHeightInPoints()
    {
        return (short)FontHeightInPoints;
    }

    public void setFontHeight(short height)
    {
        FontHeightInPoints = height / 20.0;
    }

    public void setFontHeightInPoints(short height)
    {
        FontHeightInPoints = height;
    }

    public string getFontName()
    {
        return FontName;
    }

    public void setFontName(string? name)
    {
        FontName = name ?? DEFAULT_FONT_NAME;
    }

    public bool getItalic()
    {
        return Italic;
    }

    public void setItalic(bool italic)
    {
        Italic = italic;
    }

    public bool getStrikeout()
    {
        return Strikeout;
    }

    public void setStrikeout(bool strikeout)
    {
        Strikeout = strikeout;
    }

    public byte getUnderline()
    {
        return Underline;
    }

    public void setUnderline(byte underline)
    {
        Underline = underline;
    }

    internal string GetFontHeightText()
    {
        return FontHeightInPoints.ToString("0.0################", CultureInfo.InvariantCulture);
    }
}
