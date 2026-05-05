namespace DotnetPoi.XSSF.UserModel;

/// <summary>
/// Represents a rich text string in OOXML — a sequence of runs where each run
/// may carry its own formatting (bold, italic, font, size, color, etc.).
/// Ported from org.apache.poi.xssf.usermodel.XSSFRichTextString.
/// </summary>
public sealed class XSSFRichTextString
{
    /// <summary>A single text run with optional formatting.</summary>
    public sealed class TextRun
    {
        public string Text { get; set; } = string.Empty;
        public bool Bold { get; set; }
        public bool Italic { get; set; }
        public bool Underline { get; set; }
        public bool Strikethrough { get; set; }
        /// <summary>Font size in points (e.g. 11.0). 0 means unset.</summary>
        public double FontSize { get; set; }
        /// <summary>Font name (e.g. "Arial"). Null means unset.</summary>
        public string? FontName { get; set; }
        /// <summary>Font color as hex RGB, e.g. "FF0000". Null means unset.</summary>
        public string? Color { get; set; }
    }

    private readonly List<TextRun> _runs = new();

    /// <summary>Create an empty rich text string.</summary>
    public XSSFRichTextString() { }

    /// <summary>Create a rich text string from plain text (single unformatted run).</summary>
    public XSSFRichTextString(string text)
    {
        _runs.Add(new TextRun { Text = text ?? string.Empty });
    }

    /// <summary>Create a rich text string from a pre-populated list of runs.</summary>
    public XSSFRichTextString(List<TextRun> runs)
    {
        _runs.AddRange(runs);
    }

    /// <summary>All runs in this rich text string.</summary>
    public IReadOnlyList<TextRun> Runs => _runs;

    /// <summary>The plain-text content (concatenation of all run texts).</summary>
    public string getString()
    {
        var sb = new System.Text.StringBuilder();
        foreach (var r in _runs)
            sb.Append(r.Text);
        return sb.ToString();
    }

    /// <summary>True if there is more than one run or any run carries formatting.</summary>
    public bool IsRichText
    {
        get
        {
            if (_runs.Count > 1) return true;
            if (_runs.Count == 1)
            {
                var r = _runs[0];
                return r.Bold || r.Italic || r.Underline || r.Strikethrough
                    || r.FontSize > 0 || r.FontName is not null || r.Color is not null;
            }
            return false;
        }
    }

    /// <summary>Add an unformatted text run.</summary>
    public void addRun(string text)
    {
        _runs.Add(new TextRun { Text = text ?? string.Empty });
    }

    /// <summary>Add a formatted text run.</summary>
    public void addRun(string text, bool bold, bool italic, bool underline, bool strikethrough,
        double fontSize = 0, string? fontName = null, string? color = null)
    {
        _runs.Add(new TextRun
        {
            Text = text ?? string.Empty,
            Bold = bold,
            Italic = italic,
            Underline = underline,
            Strikethrough = strikethrough,
            FontSize = fontSize,
            FontName = fontName,
            Color = color
        });
    }
}
