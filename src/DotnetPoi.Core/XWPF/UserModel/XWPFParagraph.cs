namespace DotnetPoi.XWPF.UserModel;

public sealed class XWPFParagraph
{
    private readonly List<XWPFRun> _runs = new();

    internal XWPFParagraph(XWPFDocument document)
    {
        Document = document;
    }

    internal XWPFDocument Document { get; }
    internal IReadOnlyList<XWPFRun> Runs => _runs;

    public IReadOnlyList<XWPFRun> getRuns() => _runs;

    public XWPFRun createRun()
    {
        var run = new XWPFRun(this);
        _runs.Add(run);
        return run;
    }

    public string getText()
    {
        var sb = new System.Text.StringBuilder();
        foreach (var run in _runs)
        {
            if (run.TextValue is not null)
            {
                sb.Append(run.TextValue);
            }
        }
        return sb.ToString();
    }
}
