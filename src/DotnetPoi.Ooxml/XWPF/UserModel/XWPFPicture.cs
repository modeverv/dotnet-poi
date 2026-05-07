namespace DotnetPoi.XWPF.UserModel;

public sealed class XWPFPicture
{
    // OOXML stores rotation in 60,000ths of a degree
    private const int OoxmlDegreeFactor = 60_000;

    internal XWPFPicture(XWPFPictureData data, string filename, long width, long height, string relationId, long drawingId)
    {
        PictureData = data;
        Filename = filename;
        Width = width;
        Height = height;
        RelationId = relationId;
        DrawingId = drawingId;
    }

    internal XWPFPictureData PictureData { get; }
    internal string Filename { get; }
    internal long Width { get; }
    internal long Height { get; }
    internal string RelationId { get; }
    internal long DrawingId { get; }
    internal int RotationAttribute { get; private set; }

    public double getRotation() => RotationAttribute / (double)OoxmlDegreeFactor;

    public void setRotation(double degrees)
    {
        var normalised = degrees % 360.0;
        if (normalised < 0) normalised += 360.0;
        RotationAttribute = (int)Math.Round(normalised * OoxmlDegreeFactor);
    }

    internal void SetRotationAttribute(int attribute) => RotationAttribute = attribute;
}
