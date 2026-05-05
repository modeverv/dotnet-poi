namespace DotnetPoi.XWPF.UserModel;

public sealed class XWPFRun
{
    private readonly XWPFParagraph _paragraph;
    private string? _text;
    private bool _bold;
    private bool _italic;
    private string? _fontName;
    private double _fontSize; // -1 = unset
    private string? _color;
    private bool _underline;
    private bool _strike;
    private readonly List<XWPFPicture> _pictures = new();

    internal XWPFRun(XWPFParagraph paragraph)
    {
        _paragraph = paragraph;
    }

    internal string? TextValue => _text;
    internal bool Bold => _bold;
    internal bool Italic => _italic;
    internal string? FontName => _fontName;
    internal double FontSize => _fontSize;
    internal string? Color => _color;
    internal bool Underline => _underline;
    internal bool Strike => _strike;
    internal IReadOnlyList<XWPFPicture> Pictures => _pictures;

    public IReadOnlyList<XWPFPicture> getEmbeddedPictures() => _pictures;

    public void setText(string text) => _text = text;

    public string? getText(int pos) => _text;

    public void setBold(bool bold) => _bold = bold;

    public bool isBold() => _bold;

    public void setItalic(bool italic) => _italic = italic;

    public bool isItalic() => _italic;

    public void setFontName(string fontName) => _fontName = fontName;

    public string? getFontName() => _fontName;

    public void setFontSize(double size) => _fontSize = size;

    public double getFontSize() => _fontSize;

    public void setColor(string color) => _color = color;

    public string? getColor() => _color;

    public void setUnderline(bool underline) => _underline = underline;

    public bool isUnderline() => _underline;

    public void setStrike(bool strike) => _strike = strike;

    public bool isStrike() => _strike;

    /// <summary>
    /// Adds an inline image to this run.
    /// </summary>
    /// <param name="pictureData">Raw image bytes.</param>
    /// <param name="pictureType">One of the PICTURE_TYPE_* constants on XWPFPictureData.</param>
    /// <param name="filename">Filename stored in the drawing metadata.</param>
    /// <param name="width">Width in EMU (1 inch = 914400 EMU).</param>
    /// <param name="height">Height in EMU.</param>
    public XWPFPicture addPicture(byte[] pictureData, int pictureType, string filename, int width, int height)
    {
        ArgumentNullException.ThrowIfNull(pictureData);
        var doc = _paragraph.Document;
        var data = doc.AddPictureData(pictureData, pictureType);
        // rId1 is reserved for settings.xml; images start at rId2 = rId{Index + 1}
        var relationId = $"rId{data.Index + 1}";
        var drawingId = doc.ReserveDrawingId();
        var picture = new XWPFPicture(data, filename, width, height, relationId, drawingId);
        _pictures.Add(picture);
        return picture;
    }

    internal void AttachPicture(XWPFPicture picture) => _pictures.Add(picture);
}
