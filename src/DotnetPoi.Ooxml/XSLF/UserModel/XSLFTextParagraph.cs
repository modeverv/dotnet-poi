namespace DotnetPoi.XSLF.UserModel;

/// <summary>
/// A paragraph of text inside a PPTX text box (auto shape).
/// Corresponds to an <c>a:p</c> element. Contains one or more <see cref="XSLFTextRun"/>.
/// </summary>
public sealed class XSLFTextParagraph
{
    private readonly List<XSLFTextRun> _runs = new();

    /// <summary>The text runs in this paragraph.</summary>
    public IReadOnlyList<XSLFTextRun> Runs => _runs;

    /// <summary>Creates and appends a new run with the given text, returning it.</summary>
    public XSLFTextRun addRun(string text)
    {
        var run = new XSLFTextRun(text);
        _runs.Add(run);
        return run;
    }

    /// <summary>Creates and appends an empty run, returning it.</summary>
    public XSLFTextRun addRun()
    {
        var run = new XSLFTextRun();
        _runs.Add(run);
        return run;
    }

    internal void AddRun(XSLFTextRun run) => _runs.Add(run);

    /// <summary>Plain-text concatenation of all runs.</summary>
    public string getPlainText()
    {
        return string.Concat(_runs.Select(r => r.Text));
    }
}
