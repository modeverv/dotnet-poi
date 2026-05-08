namespace DotnetPoi.XSSF.UserModel;

public sealed class XSSFPicture
{
    // OOXML stores rotation in 60,000ths of a degree (same unit as CTTransform2D.rot)
    private const int OoxmlDegreeFactor = 60_000;

    internal XSSFPicture(XSSFDrawing drawing, XSSFClientAnchor anchor, int pictureIndex, int shapeId, string relationshipId)
    {
        Drawing = drawing;
        Anchor = anchor;
        PictureIndex = pictureIndex;
        ShapeId = shapeId;
        RelationshipId = relationshipId;
    }

    internal XSSFDrawing Drawing { get; }

    internal XSSFClientAnchor Anchor { get; }

    internal int PictureIndex { get; }

    internal int ShapeId { get; }

    internal string RelationshipId { get; }

    // Rotation in 60,000ths of a degree; 0 = no rotation
    internal int RotationAttribute { get; private set; }

    /// <summary>Returns the rotation angle in degrees (clockwise, 0–360).</summary>
    public double getRotation() => RotationAttribute / (double)OoxmlDegreeFactor;

    /// <summary>Sets the rotation angle in degrees (clockwise). Value is normalised to [0, 360).</summary>
    public void setRotation(double degrees)
    {
        var normalised = degrees % 360.0;
        if (normalised < 0) normalised += 360.0;
        RotationAttribute = (int)Math.Round(normalised * OoxmlDegreeFactor);
    }

    /// <summary>Sets the rotation angle directly as the raw OOXML attribute value (60,000ths of a degree).</summary>
    internal void SetRotationAttribute(int attribute) => RotationAttribute = attribute;
}
