namespace DotnetPoi.XSLF.UserModel;

/// <summary>A single slide in a PPTX presentation.</summary>
public sealed class XSLFSlide
{
    private readonly List<XSLFPictureShape> _shapes = new();
    private int _nextShapeId = 2; // 1 is reserved for the group shape container

    internal XSLFSlide() { }

    /// <summary>All picture shapes added to this slide.</summary>
    public IReadOnlyList<XSLFPictureShape> getShapes() => _shapes;

    /// <summary>
    /// Creates a picture shape referencing the picture at the given 0-based index.
    /// The anchor defaults to (0,0) with zero size — call setAnchor() on the returned shape.
    /// </summary>
    public XSLFPictureShape createPicture(int pictureIndex)
    {
        // Caller supplies the resolved XSLFPictureData via the overload used by XMLSlideShow
        throw new InvalidOperationException("Use XMLSlideShow.createPicture(slide, pictureIndex) instead.");
    }

    internal XSLFPictureShape CreatePicture(XSLFPictureData data, string relationId)
    {
        int shapeId = _nextShapeId++;
        var shape = new XSLFPictureShape(data, relationId, shapeId);
        _shapes.Add(shape);
        return shape;
    }

    internal void AttachShape(XSLFPictureShape shape) => _shapes.Add(shape);
}
