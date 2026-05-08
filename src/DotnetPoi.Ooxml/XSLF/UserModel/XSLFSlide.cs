namespace DotnetPoi.XSLF.UserModel;

/// <summary>A single slide in a PPTX presentation.</summary>
public sealed class XSLFSlide
{
    private readonly List<XSLFPictureShape> _shapes = new();
    private readonly List<XSLFAutoShape> _autoShapes = new();
    private readonly List<XSLFTable> _tables = new();
    private int _nextShapeId = 2; // 1 is reserved for the group shape container

    internal XSLFSlide() { }

    /// <summary>All picture shapes added to this slide.</summary>
    public IReadOnlyList<XSLFPictureShape> getShapes() => _shapes;

    /// <summary>All auto shapes (text boxes) on this slide.</summary>
    public IReadOnlyList<XSLFAutoShape> getAutoShapes() => _autoShapes;

    /// <summary>All tables on this slide.</summary>
    public IReadOnlyList<XSLFTable> getTables() => _tables;

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

    /// <summary>Creates a text box (auto shape) on this slide.</summary>
    public XSLFAutoShape createTextBox()
    {
        int shapeId = _nextShapeId++;
        var shape = new XSLFAutoShape(shapeId);
        _autoShapes.Add(shape);
        return shape;
    }

    internal void AttachAutoShape(XSLFAutoShape shape) => _autoShapes.Add(shape);

    /// <summary>Creates a table shape on this slide.</summary>
    public XSLFTable createTable()
    {
        int shapeId = _nextShapeId++;
        var table = new XSLFTable(shapeId);
        _tables.Add(table);
        return table;
    }

    internal void AttachTable(XSLFTable table) => _tables.Add(table);

    // ---- Layout reference ----

    /// <summary>
    /// Relative path from this slide's rels file to its assigned slide layout.
    /// Null means use the presentation default (slideLayout1.xml).
    /// </summary>
    internal string? LayoutRelPath { get; set; }

    // ---- Raw XML preservation for unknown spTree children (group shapes, connectors, etc.) ----

    private readonly List<string> _preservedRawElements = new();

    /// <summary>
    /// Raw XML strings of unknown child elements inside p:spTree that the model does not handle
    /// (e.g. p:grpSp, p:cxnSp). These are emitted verbatim during write so they survive round-trip.
    /// </summary>
    public IReadOnlyList<string> getPreservedRawElements() => _preservedRawElements;

    internal void addPreservedRawElement(string rawXml) => _preservedRawElements.Add(rawXml);
}
