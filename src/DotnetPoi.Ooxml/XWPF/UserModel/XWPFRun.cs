namespace DotnetPoi.XWPF.UserModel;

public sealed class XWPFRun
{
    internal enum ContentChildKind
    {
        Text,
        Raw
    }

    internal sealed class ContentChild
    {
        private ContentChild(ContentChildKind kind, string? rawXml)
        {
            Kind = kind;
            RawXml = rawXml;
        }

        internal ContentChildKind Kind { get; }
        internal string? RawXml { get; }

        internal static ContentChild ForText() => new(ContentChildKind.Text, null);
        internal static ContentChild ForRaw(string rawXml) => new(ContentChildKind.Raw, rawXml);
    }

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
    private string? _hyperlinkUrl;
    private bool _hasTextContentChild;
    private readonly List<ContentChild> _contentChildren = new();
    private readonly List<string> _commentReferenceIds = new();
    internal string? HyperlinkUrl => _hyperlinkUrl;
    internal string? HyperlinkRelId { get; set; }
    internal IReadOnlyList<ContentChild> ContentChildren => _contentChildren;
    internal bool HasContent => _text is not null || _contentChildren.Count > 0;

    // Raw XML for anchored (floating) images inside this run
    private readonly List<string> _rawAnchorXml = new();
    internal IReadOnlyList<string> RawAnchorXml => _rawAnchorXml;
    internal void addRawAnchorXml(string xml) => _rawAnchorXml.Add(xml);

    // Text found in DrawingML/VML text boxes. This is extraction-only text:
    // it should be visible through getText(), but not serialized as a normal run.
    private readonly List<string> _textBoxText = new();
    internal IReadOnlyList<string> TextBoxText => _textBoxText;
    internal void addTextBoxText(string text)
    {
        if (!string.IsNullOrEmpty(text))
            _textBoxText.Add(text);
    }

    // Raw XML for unmodeled children of run properties (rPr), e.g., w:shd, highlight, etc.
    private readonly List<string> _preservedRawRPrChildren = new();
    internal IReadOnlyList<string> PreservedRawRPrChildren => _preservedRawRPrChildren;
    internal bool HasPreservedRawRPrChildren => _preservedRawRPrChildren.Count > 0;
    internal void addPreservedRawRPrChild(string rawXml) => _preservedRawRPrChildren.Add(rawXml);

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

    public void setText(string text)
    {
        _text = text;
        if (!_hasTextContentChild)
        {
            _contentChildren.Add(ContentChild.ForText());
            _hasTextContentChild = true;
        }
    }

    internal void addPreservedRawContentElement(string rawXml) =>
        AddPreservedRawContentElement(rawXml);

    private void AddPreservedRawContentElement(string rawXml)
    {
        _contentChildren.Add(ContentChild.ForRaw(rawXml));
        if (TryGetCommentReferenceId(rawXml, out var id))
            _commentReferenceIds.Add(id);
    }

    public IReadOnlyList<string> getCommentReferenceIds() => _commentReferenceIds;

    public string? getText(int pos) => _text;

    internal string getTextForExtraction()
    {
        if (_textBoxText.Count == 0)
            return _text ?? string.Empty;

        if (string.IsNullOrEmpty(_text))
            return string.Join("\n", _textBoxText);

        return _text + "\n" + string.Join("\n", _textBoxText);
    }

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

    /// <summary>Sets a hyperlink URL on this run. The run's text becomes the clickable link.</summary>
    public void setHyperlink(string url) => _hyperlinkUrl = url;

    /// <summary>Returns the hyperlink URL, or null if this run is not a hyperlink.</summary>
    public string? getHyperlink() => _hyperlinkUrl;

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
        Guard.ThrowIfNull(pictureData, nameof(pictureData));
        var doc = _paragraph.Document;
        var data = doc.AddPictureData(pictureData, pictureType);
        // rId1 is reserved for settings.xml; rId2 is for styles.xml; images start at rId{Index + 2}
        var relationId = $"rId{data.Index + 2}";
        var drawingId = doc.ReserveDrawingId();
        var picture = new XWPFPicture(data, filename, width, height, relationId, drawingId);
        _pictures.Add(picture);
        return picture;
    }

    internal void AttachPicture(XWPFPicture picture) => _pictures.Add(picture);

    private static bool TryGetCommentReferenceId(string rawXml, out string id)
    {
        id = string.Empty;
        if (!ContainsCommentReference(rawXml))
            return false;

        foreach (var marker in new[] { " w:id=", " id=" })
        {
            var index = rawXml.IndexOf(marker, StringComparison.Ordinal);
            if (index < 0)
                continue;

            var valueStart = index + marker.Length;
            if (valueStart >= rawXml.Length)
                continue;

            var quote = rawXml[valueStart];
            if (quote != '"' && quote != '\'')
                continue;

            var valueEnd = rawXml.IndexOf(quote, valueStart + 1);
            if (valueEnd <= valueStart)
                continue;

            id = rawXml.Substring(valueStart + 1, valueEnd - valueStart - 1);
            return true;
        }

        return false;
    }

    private static bool ContainsCommentReference(string rawXml) =>
        rawXml.StartsWith("<w:commentReference", StringComparison.Ordinal)
        || rawXml.StartsWith("<commentReference", StringComparison.Ordinal)
        || rawXml.IndexOf("<w:commentReference ", StringComparison.Ordinal) >= 0
        || rawXml.IndexOf("<commentReference ", StringComparison.Ordinal) >= 0
        || rawXml.IndexOf("<w:commentReference/>", StringComparison.Ordinal) >= 0
        || rawXml.IndexOf("<commentReference/>", StringComparison.Ordinal) >= 0;
}
