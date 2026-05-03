namespace DotnetPoi.XSSF.UserModel;

public sealed class XSSFPicture
{
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
}
