namespace DotnetPoi.XSSF.UserModel;

public sealed class XSSFDrawing
{
    private readonly List<XSSFPicture> _pictures = new();
    private readonly XSSFSheet _sheet;

    internal XSSFDrawing(XSSFSheet sheet, int drawingIndex)
    {
        _sheet = sheet;
        DrawingIndex = drawingIndex;
    }

    internal int DrawingIndex { get; }

    internal IReadOnlyList<XSSFPicture> Pictures => _pictures;

    public XSSFClientAnchor createAnchor(int dx1, int dy1, int dx2, int dy2, int col1, int row1, int col2, int row2)
    {
        return new XSSFClientAnchor(dx1, dy1, dx2, dy2, col1, row1, col2, row2);
    }

    public XSSFPicture createPicture(XSSFClientAnchor anchor, int pictureIndex)
    {
        ArgumentNullException.ThrowIfNull(anchor);
        _sheet.getWorkbook().GetPictureData(pictureIndex);

        var relationshipId = "rId" + (_pictures.Count + 1).ToString(System.Globalization.CultureInfo.InvariantCulture);
        var picture = new XSSFPicture(this, anchor, pictureIndex, _pictures.Count + 1, relationshipId);
        _pictures.Add(picture);
        return picture;
    }

    public IReadOnlyList<XSSFPicture> getShapes()
    {
        return _pictures;
    }
}
