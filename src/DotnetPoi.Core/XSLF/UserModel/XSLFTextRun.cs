namespace DotnetPoi.XSLF.UserModel;

/// <summary>
/// A single text run inside a PPTX paragraph.
/// Corresponds to an <c>a:r</c> element in the OOXML.
/// </summary>
public sealed class XSLFTextRun
{
    /// <summary>The text content of this run.</summary>
    public string Text { get; set; } = string.Empty;

    /// <summary>Whether bold is applied.</summary>
    public bool Bold { get; set; }

    /// <summary>Whether italic is applied.</summary>
    public bool Italic { get; set; }

    /// <summary>Whether underline is applied.</summary>
    public bool Underline { get; set; }

    /// <summary>Whether strikethrough is applied.</summary>
    public bool Strikethrough { get; set; }

    /// <summary>Font size in points. 0 means unspecified (use theme default).</summary>
    public double FontSize { get; set; }

    /// <summary>Font family, e.g. "Calibri". Null means unspecified.</summary>
    public string? FontName { get; set; }

    /// <summary>Color in ARGB hex e.g. "FF0000" for red. Null means unspecified.</summary>
    public string? Color { get; set; }

    public XSLFTextRun() { }

    public XSLFTextRun(string text)
    {
        Text = text;
    }
}
