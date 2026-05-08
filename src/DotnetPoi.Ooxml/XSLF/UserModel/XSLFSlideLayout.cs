namespace DotnetPoi.XSLF.UserModel;

/// <summary>
/// Represents a slide layout within a slide master.
/// Ported from org.apache.poi.xslf.usermodel.XSLFSlideLayout.
/// </summary>
public sealed class XSLFSlideLayout
{
    /// <summary>Display name of this layout (e.g. "Title Slide", "Title and Content").</summary>
    public string Name { get; }

    /// <summary>OOXML layout type attribute (e.g. "title", "obj", "blank", "titleOnly").</summary>
    public string Type { get; }

    /// <summary>Full zip-relative path, e.g. "ppt/slideLayouts/slideLayout1.xml".</summary>
    internal string ZipPath { get; }

    /// <summary>Relative path from a slide rels file, e.g. "../slideLayouts/slideLayout1.xml".</summary>
    internal string SlideRelPath => "../" + ZipPath.Substring("ppt/".Length);

    internal XSLFSlideLayout(string name, string type, string zipPath)
    {
        Name = name;
        Type = type;
        ZipPath = zipPath;
    }

    public override string ToString() => $"{Name} ({Type})";
}
