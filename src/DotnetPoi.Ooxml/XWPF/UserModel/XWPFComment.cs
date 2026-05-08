using System.Xml.Linq;

namespace DotnetPoi.XWPF.UserModel;

/// <summary>Representation of a DOCX comment from word/comments.xml.</summary>
public sealed class XWPFComment
{
    private readonly Action? _onChanged;

    internal XWPFComment(string id, string? author, string? initials, string? date, string text, Action? onChanged = null)
    {
        Id = id;
        Author = author;
        Initials = initials;
        Date = date;
        Text = text;
        _onChanged = onChanged;
    }

    public string Id { get; }
    public string? Author { get; private set; }
    public string? Initials { get; private set; }
    public string? Date { get; private set; }
    public string Text { get; private set; }

    public string getId() => Id;
    public string? getAuthor() => Author;
    public string? getInitials() => Initials;
    public string? getDate() => Date;
    public string getText() => Text;

    public void setAuthor(string? author)
    {
        Author = author;
        _onChanged?.Invoke();
    }

    public void setInitials(string? initials)
    {
        Initials = initials;
        _onChanged?.Invoke();
    }

    public void setDate(string? date)
    {
        Date = date;
        _onChanged?.Invoke();
    }

    public void setText(string text)
    {
        Text = text;
        _onChanged?.Invoke();
    }

    internal static XWPFComment FromXml(XElement commentElement, XNamespace wordNamespace, Action? onChanged = null)
    {
        var id = (string?)commentElement.Attribute(wordNamespace + "id") ?? "-1";
        var author = (string?)commentElement.Attribute(wordNamespace + "author");
        var initials = (string?)commentElement.Attribute(wordNamespace + "initials");
        var date = (string?)commentElement.Attribute(wordNamespace + "date");

        var paragraphTexts = commentElement
            .Elements(wordNamespace + "p")
            .Select(p => string.Concat(p.Descendants(wordNamespace + "t").Select(t => (string?)t ?? string.Empty)));

        return new XWPFComment(id, author, initials, date, string.Join("\n", paragraphTexts), onChanged);
    }
}
