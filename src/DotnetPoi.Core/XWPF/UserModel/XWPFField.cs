namespace DotnetPoi.XWPF.UserModel;

/// <summary>
/// Represents a Word field (TOC, PAGE, MERGEFIELD, etc.) in a paragraph.
/// Fields are serialized as a sequence of <c>fldChar</c> / <c>instrText</c> runs
/// in the OOXML WordprocessingML format.
/// </summary>
public sealed class XWPFField
{
    /// <summary>
    /// The field instruction text, e.g. "TOC \\o \"1-3\" \\h" or "PAGE" or "MERGEFIELD Name".
    /// </summary>
    public string Instruction { get; set; } = string.Empty;

    /// <summary>
    /// The field result text — the displayed value (e.g. "1" for a page number).
    /// May be empty for fields that have not been evaluated.
    /// </summary>
    public string Result { get; set; } = string.Empty;

    /// <summary>
    /// Creates a new field with the given instruction and (optional) result.
    /// </summary>
    public XWPFField() { }

    /// <summary>
    /// Creates a new field with the given instruction and (optional) result.
    /// </summary>
    public XWPFField(string instruction, string result = "")
    {
        Instruction = instruction;
        Result = result;
    }
}
