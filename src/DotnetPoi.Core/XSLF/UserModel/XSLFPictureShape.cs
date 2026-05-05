namespace DotnetPoi.XSLF.UserModel;

/// <summary>
/// A picture shape on a PPTX slide.
/// Position and size are stored in EMU (English Metric Units).
/// 1 inch = 914400 EMU. Rotation is stored in 60,000ths of a degree, matching OOXML CTTransform2D.rot.
/// </summary>
public sealed class XSLFPictureShape
{
    private const int OoxmlDegreeFactor = 60_000;

    internal XSLFPictureShape(XSLFPictureData pictureData, string relationId, int shapeId)
    {
        PictureData = pictureData;
        RelationId  = relationId;
        ShapeId     = shapeId;
    }

    internal XSLFPictureData PictureData       { get; }
    internal string          RelationId        { get; }
    internal int             ShapeId           { get; }
    internal long            AnchorX           { get; private set; }
    internal long            AnchorY           { get; private set; }
    internal long            AnchorCx          { get; private set; }
    internal long            AnchorCy          { get; private set; }
    internal int             RotationAttribute { get; private set; }
    internal bool            FlipH             { get; private set; }
    internal bool            FlipV             { get; private set; }

    /// <summary>Sets position and size in EMU.</summary>
    public void setAnchor(long x, long y, long cx, long cy)
    {
        AnchorX  = x;
        AnchorY  = y;
        AnchorCx = cx;
        AnchorCy = cy;
    }

    /// <summary>Returns rotation in degrees.</summary>
    public double getRotation() => RotationAttribute / (double)OoxmlDegreeFactor;

    /// <summary>Sets rotation in degrees (clockwise, normalised to [0, 360)).</summary>
    public void setRotation(double degrees)
    {
        var normalised = degrees % 360.0;
        if (normalised < 0) normalised += 360.0;
        RotationAttribute = (int)Math.Round(normalised * OoxmlDegreeFactor);
    }

    public void setFlipHorizontal(bool flip) => FlipH = flip;
    public bool getFlipHorizontal()          => FlipH;

    public void setFlipVertical(bool flip)   => FlipV = flip;
    public bool getFlipVertical()            => FlipV;

    public long getAnchorX()  => AnchorX;
    public long getAnchorY()  => AnchorY;
    public long getAnchorCx() => AnchorCx;
    public long getAnchorCy() => AnchorCy;

    public XSLFPictureData getPictureData()    => PictureData;

    /// <summary>Returns the raw OOXML rotation value (60,000ths of a degree).</summary>
    public int getRotationAttribute()          => RotationAttribute;

    internal void SetRotationAttribute(int attribute) => RotationAttribute = attribute;
}
