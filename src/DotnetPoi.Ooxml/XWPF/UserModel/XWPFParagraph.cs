namespace DotnetPoi.XWPF.UserModel;

/// <summary>OOXML paragraph alignment (w:jc val attribute).</summary>
public enum ParagraphAlignment
{
    Left,
    Center,
    Right,
    Both
}

/// <summary>Line spacing rule.</summary>
public enum LineSpacingRule
{
    Auto,
    AtLeast,
    Exact
}

public sealed class XWPFParagraph
{
    internal enum ChildKind
    {
        Run,
        Field,
        Raw
    }

    internal sealed class Child
    {
        private Child(ChildKind kind, XWPFRun? run, XWPFField? field, string? rawXml)
        {
            Kind = kind;
            Run = run;
            Field = field;
            RawXml = rawXml;
        }

        internal ChildKind Kind { get; }
        internal XWPFRun? Run { get; }
        internal XWPFField? Field { get; }
        internal string? RawXml { get; }

        internal static Child ForRun(XWPFRun run) => new(ChildKind.Run, run, null, null);
        internal static Child ForField(XWPFField field) => new(ChildKind.Field, null, field, null);
        internal static Child ForRaw(string rawXml) => new(ChildKind.Raw, null, null, rawXml);
    }

    private readonly List<XWPFRun> _runs = new();
    private readonly List<XWPFField> _fields = new();
    private readonly List<string> _preservedRawElements = new();
    private readonly List<Child> _children = new();
    internal IReadOnlyList<string> PreservedRawElements => _preservedRawElements;
    internal IReadOnlyList<Child> Children => _children;
    internal void addPreservedRawElement(string rawXml)
    {
        _preservedRawElements.Add(rawXml);
        _children.Add(Child.ForRaw(rawXml));
    }
    // Paragraph-level section properties (sectPr inside pPr for section breaks)
    private string? _preservedSectPr;
    internal string? PreservedSectPr => _preservedSectPr;
    internal void setPreservedSectPr(string rawXml) => _preservedSectPr = rawXml;
    // Raw XML for unmodeled children of paragraph-level run properties (pPr/rPr), e.g., w:shd at paragraph level
    private readonly List<string> _preservedRawPPrRPrChildren = new();
    internal IReadOnlyList<string> PreservedRawPPrRPrChildren => _preservedRawPPrRPrChildren;
    internal bool HasPreservedRawPPrRPrChildren => _preservedRawPPrRPrChildren.Count > 0;
    internal void addPreservedRawPPrRPrChild(string rawXml) => _preservedRawPPrRPrChildren.Add(rawXml);
    private ParagraphAlignment? _alignment;
    // Indentation (twips = 1/1440 inch)
    private int _indentLeft;
    private int _indentRight;
    private int _indentFirstLine;
    private int _indentHanging;
    // Spacing (twips = 1/1440 inch)
    private int _spacingBefore;
    private int _spacingAfter;
    private int _spacingBetween; // line spacing, in 240ths of a line
    private LineSpacingRule _lineRule = LineSpacingRule.Auto;
    // Numbering
    private int? _numId;
    private int _ilvl;
    private string? _styleId;

    internal XWPFParagraph(XWPFDocument document)
    {
        Document = document;
    }

    internal XWPFDocument Document { get; }
    internal IReadOnlyList<XWPFRun> Runs => _runs;
    internal IReadOnlyList<XWPFField> Fields => _fields;
    internal ParagraphAlignment? Alignment => _alignment;
    internal int IndentLeft => _indentLeft;
    internal int IndentRight => _indentRight;
    internal int IndentFirstLine => _indentFirstLine;
    internal int IndentHanging => _indentHanging;
    internal int SpacingBefore => _spacingBefore;
    internal int SpacingAfter => _spacingAfter;
    internal int SpacingBetween => _spacingBetween;
    internal LineSpacingRule LineRule => _lineRule;
    internal int? NumId => _numId;
    internal int Ilvl => _ilvl;

    public IReadOnlyList<XWPFRun> getRuns() => _runs;

    public XWPFRun createRun()
    {
        var run = new XWPFRun(this);
        _runs.Add(run);
        _children.Add(Child.ForRun(run));
        return run;
    }

    public IReadOnlyList<XWPFField> getFields() => _fields;

    /// <summary>
    /// Adds a field (e.g. TOC, PAGE, MERGEFIELD) to this paragraph.
    /// Fields are serialized as fldChar/instrText sequences in the OOXML output.
    /// </summary>
    /// <param name="instruction">The field instruction text, e.g. "TOC \\o \"1-3\"" or "PAGE".</param>
    /// <param name="result">The field result (displayed text), if available.</param>
    /// <returns>The newly created <see cref="XWPFField"/>.</returns>
    public XWPFField addField(string instruction, string result = "")
    {
        var field = new XWPFField(instruction, result);
        _fields.Add(field);
        _children.Add(Child.ForField(field));
        return field;
    }

    public string getText()
    {
        var sb = new System.Text.StringBuilder();
        foreach (var run in _runs)
        {
            sb.Append(run.getTextForExtraction());
        }
        return sb.ToString();
    }

    public void setAlignment(ParagraphAlignment align) => _alignment = align;
    public ParagraphAlignment? getAlignment() => _alignment;

    public void setIndentationLeft(int twips) => _indentLeft = twips;
    public int getIndentationLeft() => _indentLeft;
    public void setIndentationRight(int twips) => _indentRight = twips;
    public int getIndentationRight() => _indentRight;
    public void setIndentationFirstLine(int twips) => _indentFirstLine = twips;
    public int getIndentationFirstLine() => _indentFirstLine;
    public void setIndentationHanging(int twips) => _indentHanging = twips;
    public int getIndentationHanging() => _indentHanging;

    public void setSpacingBefore(int twips) => _spacingBefore = twips;
    public int getSpacingBefore() => _spacingBefore;
    public void setSpacingAfter(int twips) => _spacingAfter = twips;
    public int getSpacingAfter() => _spacingAfter;
    public void setSpacingBetween(int twips) => _spacingBetween = twips;
    public int getSpacingBetween() => _spacingBetween;
    public void setLineSpacingRule(LineSpacingRule rule) => _lineRule = rule;
    public LineSpacingRule getLineSpacingRule() => _lineRule;

    public void setNumId(int? numId) => _numId = numId;
    public int? getNumId() => _numId;
    public void setIlvl(int ilvl) => _ilvl = ilvl;
    public int getIlvl() => _ilvl;

    public string? getStyleID() => _styleId;
    public void setStyle(string? styleId) => _styleId = styleId;

    /// <summary>Convenience: set up a bullet list numbering.</summary>
    public void setBulletList()
    {
        _numId = Document.GetOrCreateNumbering(XWPFDocument.NumberingFormat.Bullet);
        _ilvl = 0;
    }

    /// <summary>Convenience: set up a numbered list.</summary>
    public void setNumberedList()
    {
        _numId = Document.GetOrCreateNumbering(XWPFDocument.NumberingFormat.Decimal);
        _ilvl = 0;
    }
}
