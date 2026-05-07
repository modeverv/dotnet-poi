namespace DotnetPoi.XSLF.UserModel;

/// <summary>
/// A non-picture shape on a PPTX slide (typically a text box).
/// Corresponds to a <c>p:sp</c> element in the OOXML.
/// </summary>
public sealed class XSLFAutoShape
{
    private readonly List<XSLFTextParagraph> _paragraphs = new();

    internal XSLFAutoShape(int shapeId)
    {
        ShapeId = shapeId;
    }

    /// <summary>Unique shape ID within the slide.</summary>
    public int ShapeId { get; }

    /// <summary>Paragraphs of text in this shape.</summary>
    public IReadOnlyList<XSLFTextParagraph> Paragraphs => _paragraphs;

    /// <summary>Creates and appends a new paragraph, returning it.</summary>
    public XSLFTextParagraph addParagraph()
    {
        var p = new XSLFTextParagraph();
        _paragraphs.Add(p);
        return p;
    }

    internal void AddParagraph(XSLFTextParagraph p) => _paragraphs.Add(p);

    // --- anchor / position ---

    internal long AnchorX { get; private set; }
    internal long AnchorY { get; private set; }
    internal long AnchorCx { get; private set; }
    internal long AnchorCy { get; private set; }

    /// <summary>Sets position and size in EMU.</summary>
    public void setAnchor(long x, long y, long cx, long cy)
    {
        AnchorX = x;
        AnchorY = y;
        AnchorCx = cx;
        AnchorCy = cy;
    }

    public long getAnchorX() => AnchorX;
    public long getAnchorY() => AnchorY;
    public long getAnchorCx() => AnchorCx;
    public long getAnchorCy() => AnchorCy;

    // --- rotation / flip ---

    internal int RotationAttribute { get; private set; }
    internal bool FlipH { get; private set; }
    internal bool FlipV { get; private set; }

    private const int OoxmlDegreeFactor = 60_000;

    public double getRotation() => RotationAttribute / (double)OoxmlDegreeFactor;

    public void setRotation(double degrees)
    {
        var normalised = degrees % 360.0;
        if (normalised < 0) normalised += 360.0;
        RotationAttribute = (int)Math.Round(normalised * OoxmlDegreeFactor);
    }

    public void setFlipHorizontal(bool flip) => FlipH = flip;
    public bool getFlipHorizontal() => FlipH;
    public void setFlipVertical(bool flip) => FlipV = flip;
    public bool getFlipVertical() => FlipV;

    internal void SetRotationAttribute(int attribute) => RotationAttribute = attribute;
}
