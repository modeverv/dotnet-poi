using System.Globalization;
using System.IO.Compression;
using System.Text;
using System.Xml;
using DotnetPoi.SS.Xml;

namespace DotnetPoi.XWPF.UserModel;

public sealed class XWPFDocument : IDisposable
{
    private enum BodyChildKind
    {
        Paragraph,
        Table,
        Raw
    }

    private sealed class BodyChild
    {
        private BodyChild(BodyChildKind kind, XWPFParagraph? paragraph, XWPFTable? table, string? rawXml)
        {
            Kind = kind;
            Paragraph = paragraph;
            Table = table;
            RawXml = rawXml;
        }

        internal BodyChildKind Kind { get; }
        internal XWPFParagraph? Paragraph { get; }
        internal XWPFTable? Table { get; }
        internal string? RawXml { get; }

        internal static BodyChild ForParagraph(XWPFParagraph paragraph) =>
            new(BodyChildKind.Paragraph, paragraph, null, null);

        internal static BodyChild ForTable(XWPFTable table) =>
            new(BodyChildKind.Table, null, table, null);

        internal static BodyChild ForRaw(string rawXml) =>
            new(BodyChildKind.Raw, null, null, rawXml);
    }

    private sealed record PreservedContentTypeDefault(string Extension, string ContentType);
    private sealed record PreservedContentTypeOverride(string PartName, string ContentType);
    private sealed record PreservedRelationship(string Id, string Type, string Target, string? TargetMode);

    private const string NsW = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private const string NsR = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private const string NsWp = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
    private const string NsA = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private const string NsPic = "http://schemas.openxmlformats.org/drawingml/2006/picture";

    private const string RelTypeSettings = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings";
    private const string ContentTypeSettings = "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml";
    private const string ContentTypeDocx = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";
    private const string ContentTypeDocm = "application/vnd.ms-word.document.macroEnabled.main+xml";
    private const string ContentTypeVbaProject = "application/vnd.ms-office.vbaProject";
    private const string ContentTypeVbaData = "application/vnd.ms-word.vbaData+xml";
    private const string RelTypeVbaProject = "http://schemas.microsoft.com/office/2006/relationships/vbaProject";

    private const string RelTypeNumbering = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering";
    private const string ContentTypeNumbering = "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml";
    private const string RelTypeHyperlink = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
    private const string RelTypeHeader = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header";
    private const string ContentTypeHeader = "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml";
    private const string RelTypeFooter = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer";
    private const string ContentTypeFooter = "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml";

    private const string RelTypeStyles = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";
    private const string ContentTypeStyles = "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml";
    private const string RelTypeComments = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments";
    private const string ContentTypeComments = "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml";

    /// <summary>True if this document was loaded from docm (has a vbaProject.bin).</summary>
    public bool HasMacros => _vbaProjectBin != null;

    public bool isMacroEnabled() => HasMacros;

    // rId1 is always reserved for settings.xml; rId2 is for styles.xml; image rIds start at rId{Index + 2}.
    private const int ImageRelIdOffset = 2;

    private readonly List<XWPFParagraph> _paragraphs = new();
    private readonly List<XWPFComment> _comments = new();
    private readonly List<XWPFPictureData> _pictures = new();
    private readonly List<XWPFTable> _tables = new();
    private readonly List<BodyChild> _bodyChildren = new();
    private readonly List<string> _hyperlinkUrls = new();
    private long _nextDrawingId = 1;

    // Numbering support
    private readonly List<(int abstractNumId, string numFmt, string lvlText, int startVal)> _abstractNums = new();
    private readonly List<(int numId, int abstractNumId)> _numInstances = new();
    private int _nextAbstractNumId;
    private int _nextNumId = 1;

    // docm support: opaque VBA binaries preserved byte-for-byte.
    private byte[]? _vbaProjectBin;
    private byte[]? _vbaDataXml;
    private byte[]? _vbaProjectBinRels;

    // Page setup (twips = 1/1440 inch)
    private long _pageWidth = 12240;  // default: 8.5in
    private long _pageHeight = 15840; // default: 11in
    private string? _pageOrientation; // null = portrait, "landscape" = landscape
    private long _marginTop = 1440;    // 1in
    private long _marginRight = 1440;
    private long _marginBottom = 1440;
    private long _marginLeft = 1440;
    private long _marginHeader = 720;  // 0.5in
    private long _marginFooter = 720;

    // Headers / footers (three variants per section: default, first page, even page)
    private string? _headerText;         // default header
    private string? _headerFirstText;    // first page header
    private string? _headerEvenText;     // even page header
    private string? _footerText;         // default footer
    private string? _footerFirstText;    // first page footer
    private string? _footerEvenText;     // even page footer
    private bool _headerModified;
    private bool _footerModified;

    // Computed header/footer counts (how many distinct files to write)
    private int _headerCount => (!string.IsNullOrEmpty(_headerText) ? 1 : 0)
                              + (!string.IsNullOrEmpty(_headerFirstText) ? 1 : 0)
                              + (!string.IsNullOrEmpty(_headerEvenText) ? 1 : 0);
    private int _footerCount => (!string.IsNullOrEmpty(_footerText) ? 1 : 0)
                              + (!string.IsNullOrEmpty(_footerFirstText) ? 1 : 0)
                              + (!string.IsNullOrEmpty(_footerEvenText) ? 1 : 0);

    // Unknown part preservation: ZIP entries not understood by the model are stored
    // byte-for-byte and re-emitted during write, ensuring layouts, themes, styles,
    // docProps, and other unmodeled parts survive round-trip.
    private readonly Dictionary<string, byte[]> _preservedEntries
        = new(StringComparer.OrdinalIgnoreCase);
    private readonly List<PreservedContentTypeDefault> _preservedContentTypeDefaults = new();
    private readonly List<PreservedContentTypeOverride> _preservedContentTypeOverrides = new();
    private readonly List<PreservedRelationship> _preservedDocumentRelationships = new();

    // Raw XML preservation for unknown direct children of w:body in document.xml
    // (SDT / content controls, altChunk, bookmarks, track changes, etc.).
    internal readonly List<string> _preservedRawBodyElements = new();

    // Columns (newspaper columns in a section)
    private int _columnCount = 1;
    private long _columnSpacing; // twips (1/1440 inch) between columns

    // Raw XML preservation for unknown children of the final w:sectPr
    // (pgBorders, lnNumType, docGrid, formProt, etc.)
    private readonly List<string> _preservedRawSectPrChildren = new();

    // Body-level section breaks — mid-document w:sectPr elements that appear between
    // body children (paragraphs, tables). Captured as raw XML for verbatim re-emission.
    // The LAST element in this list is the final section; its children are parsed for
    // model state (pgSz, pgMar, cols, etc.) via ParseFinalSectPr().
    private readonly List<string> _preservedBodySectPr = new();

    private XWPFStyles? _styles;
    private bool _commentsModified;

    public XWPFDocument() { }

    public XWPFDocument(Stream stream)
    {
        Guard.ThrowIfNull(stream, nameof(stream));
        Load(stream);
    }

    public XWPFParagraph createParagraph()
    {
        var paragraph = new XWPFParagraph(this);
        _paragraphs.Add(paragraph);
        _bodyChildren.Add(BodyChild.ForParagraph(paragraph));
        return paragraph;
    }

    public IReadOnlyList<XWPFParagraph> getParagraphs() => _paragraphs;

    public IReadOnlyList<XWPFComment> getComments() => _comments;

    public XWPFComment? getCommentByID(string id) =>
        _comments.FirstOrDefault(comment => comment.getId() == id);

    public XWPFComment createComment(string? author = null, string? text = null, string? initials = null, string? date = null)
    {
        var comment = new XWPFComment(GetNextCommentId(), author, initials, date, text ?? string.Empty, MarkCommentsModified);
        _comments.Add(comment);
        MarkCommentsModified();
        return comment;
    }

    public bool removeComment(string id)
    {
        var index = _comments.FindIndex(comment => comment.getId() == id);
        if (index < 0)
            return false;

        _comments.RemoveAt(index);
        MarkCommentsModified();
        return true;
    }

    internal void MarkCommentsModified() => _commentsModified = true;

    private string GetNextCommentId()
    {
        var used = new HashSet<string>(_comments.Select(comment => comment.getId()), StringComparer.Ordinal);
        var id = 0;
        while (used.Contains(id.ToString(CultureInfo.InvariantCulture)))
            id++;
        return id.ToString(CultureInfo.InvariantCulture);
    }

    public XWPFTable createTable()
    {
        var table = new XWPFTable(this);
        _tables.Add(table);
        _bodyChildren.Add(BodyChild.ForTable(table));
        return table;
    }

    public IReadOnlyList<XWPFTable> getTables() => _tables;

    public XWPFStyles? getStyles() => _styles;

    public IReadOnlyList<XWPFPictureData> getAllPictures() => _pictures;

    /// <summary>Returns 0-based index of added picture data.</summary>
    public int addPicture(byte[] pictureData, int format)
    {
        Guard.ThrowIfNull(pictureData, nameof(pictureData));
        var data = AddPictureData(pictureData, format);
        return data.Index - 1;
    }

    /// <summary>Set page size in twips (1/1440 inch). Default: 12240 x 15840 (Letter).</summary>
    public void setPageSize(long widthTwips, long heightTwips)
    {
        _pageWidth = widthTwips;
        _pageHeight = heightTwips;
    }

    /// <summary>Get page width in twips.</summary>
    public long getPageWidth() => _pageWidth;

    /// <summary>Get page height in twips.</summary>
    public long getPageHeight() => _pageHeight;

    /// <summary>Set landscape orientation (default is portrait).</summary>
    public void setLandscape(bool landscape) => _pageOrientation = landscape ? "landscape" : null;

    /// <summary>True if page orientation is landscape.</summary>
    public bool isLandscape() => _pageOrientation == "landscape";

    /// <summary>Set margin in twips (1/1440 inch). Pass 0 for none.</summary>
    public void setMargins(long top, long right, long bottom, long left)
    {
        _marginTop = top;
        _marginRight = right;
        _marginBottom = bottom;
        _marginLeft = left;
    }

    public long getMarginTop() => _marginTop;
    public long getMarginRight() => _marginRight;
    public long getMarginBottom() => _marginBottom;
    public long getMarginLeft() => _marginLeft;

    /// <summary>Set newspaper columns for the final section.</summary>
    /// <param name="count">Number of columns (1 = no columns).</param>
    /// <param name="spacingTwips">Spacing between columns in twips (1/1440 inch).</param>
    public void setColumns(int count, long spacingTwips = 0)
    {
        _columnCount = count <= 0 ? 1 : count;
        _columnSpacing = spacingTwips;
    }

    /// <summary>Get number of columns in the final section.</summary>
    public int getColumnCount() => _columnCount;

    /// <summary>Get spacing between columns in twips.</summary>
    public long getColumnSpacing() => _columnSpacing;

    /// <summary>Set default header text (plain text, first section).</summary>
    public void setHeaderText(string text)
    {
        _headerText = text;
        _headerModified = true;
    }

    /// <summary>Get default header text.</summary>
    public string? getHeaderText() => _headerText;

    /// <summary>Set first-page header text.</summary>
    public void setFirstHeaderText(string text)
    {
        _headerFirstText = text;
        _headerModified = true;
    }

    /// <summary>Get first-page header text.</summary>
    public string? getFirstHeaderText() => _headerFirstText;

    /// <summary>Set even-page header text.</summary>
    public void setEvenHeaderText(string text)
    {
        _headerEvenText = text;
        _headerModified = true;
    }

    /// <summary>Get even-page header text.</summary>
    public string? getEvenHeaderText() => _headerEvenText;

    /// <summary>Set default footer text (plain text, first section).</summary>
    public void setFooterText(string text)
    {
        _footerText = text;
        _footerModified = true;
    }

    /// <summary>Get default footer text.</summary>
    public string? getFooterText() => _footerText;

    /// <summary>Set first-page footer text.</summary>
    public void setFirstFooterText(string text)
    {
        _footerFirstText = text;
        _footerModified = true;
    }

    /// <summary>Get first-page footer text.</summary>
    public string? getFirstFooterText() => _footerFirstText;

    /// <summary>Set even-page footer text.</summary>
    public void setEvenFooterText(string text)
    {
        _footerEvenText = text;
        _footerModified = true;
    }

    /// <summary>Get even-page footer text.</summary>
    public string? getEvenFooterText() => _footerEvenText;

    internal bool HasHeader => _headerCount > 0;
    internal bool HasFooter => _footerCount > 0;
    internal int HeaderCount => _headerCount;
    internal int FooterCount => _footerCount;

    private void AddPreservedRawBodyElement(string rawXml)
    {
        _preservedRawBodyElements.Add(rawXml);
        _bodyChildren.Add(BodyChild.ForRaw(rawXml));
    }

    public void write(Stream stream)
    {
        Guard.ThrowIfNull(stream, nameof(stream));
        using var archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true);
        CollectHyperlinks();

        // Emit preserved (unknown) entries first, then model entries overwrite below
        foreach (var kv in _preservedEntries)
        {
            if (_commentsModified && kv.Key.Equals("word/comments.xml", StringComparison.OrdinalIgnoreCase))
                continue;
            WriteBinaryEntry(archive, kv.Key, kv.Value);
        }

        WriteEntry(archive, "[Content_Types].xml", WriteContentTypes);
        WriteEntry(archive, "_rels/.rels", WriteRootRelationships);
        WriteEntry(archive, "word/document.xml", WriteDocument);
        WriteEntry(archive, "word/_rels/document.xml.rels", WriteDocumentRelationships);
        WriteEntry(archive, "word/settings.xml", WriteSettings);
        if (_numInstances.Count > 0)
            WriteEntry(archive, "word/numbering.xml", WriteNumbering);
        WriteEntry(archive, "word/styles.xml", WriteStyles);
        if (_commentsModified)
            WriteEntry(archive, "word/comments.xml", WriteComments);
        // headers and footers — write model content only if modified via API, otherwise preserved bytes are used
        if (_headerModified)
        {
            int hIdx = 1;
            if (!string.IsNullOrEmpty(_headerText))
            {
                var text = _headerText!; // capture for closure
                WriteEntry(archive, $"word/header{hIdx++}.xml", w => WriteHeader(w, text));
            }
            if (!string.IsNullOrEmpty(_headerFirstText))
            {
                var text = _headerFirstText!;
                WriteEntry(archive, $"word/header{hIdx++}.xml", w => WriteHeader(w, text));
            }
            if (!string.IsNullOrEmpty(_headerEvenText))
            {
                var text = _headerEvenText!;
                WriteEntry(archive, $"word/header{hIdx++}.xml", w => WriteHeader(w, text));
            }
        }
        if (_footerModified)
        {
            int fIdx = 1;
            if (!string.IsNullOrEmpty(_footerText))
            {
                var text = _footerText!;
                WriteEntry(archive, $"word/footer{fIdx++}.xml", w => WriteFooter(w, text));
            }
            if (!string.IsNullOrEmpty(_footerFirstText))
            {
                var text = _footerFirstText!;
                WriteEntry(archive, $"word/footer{fIdx++}.xml", w => WriteFooter(w, text));
            }
            if (!string.IsNullOrEmpty(_footerEvenText))
            {
                var text = _footerEvenText!;
                WriteEntry(archive, $"word/footer{fIdx++}.xml", w => WriteFooter(w, text));
            }
        }
        foreach (var pic in _pictures)
        {
            WriteBinaryEntry(archive, $"word/media/{pic.getFileName()}", pic.Data);
        }

        // docm: write VBA binaries back verbatim
        if (_vbaProjectBin != null)
        {
            WriteBinaryEntry(archive, "word/vbaProject.bin", _vbaProjectBin);
            if (_vbaDataXml != null)
                WriteBinaryEntry(archive, "word/vbaData.xml", _vbaDataXml);
            if (_vbaProjectBinRels != null)
                WriteBinaryEntry(archive, "word/_rels/vbaProject.bin.rels", _vbaProjectBinRels);
        }
    }

    public void close() { }

    public void Dispose() => close();

    /// <summary>Numbering formats supported by GetOrCreateNumbering.</summary>
    public enum NumberingFormat
    {
        Bullet,
        Decimal
    }

    /// <summary>
    /// Gets or creates a numbering (num) instance with the given format.
    /// Returns the numId to use in a paragraph's numPr.
    /// </summary>
    public int GetOrCreateNumbering(NumberingFormat format)
    {
        var (numFmt, lvlText, startVal) = format switch
        {
            NumberingFormat.Bullet => ("bullet", "\u2022", 1),
            NumberingFormat.Decimal => ("decimal", "%1", 1),
            _ => ("decimal", "%1", 1)
        };

        // Check if we already have an abstractNum with this format
        var existingAbs = _abstractNums.FirstOrDefault(a => a.numFmt == numFmt && a.lvlText == lvlText);
        int abstractNumId;
        if (existingAbs.abstractNumId == 0 && existingAbs.numFmt is null)
        {
            abstractNumId = _nextAbstractNumId++;
            _abstractNums.Add((abstractNumId, numFmt, lvlText, startVal));
        }
        else
        {
            abstractNumId = existingAbs.abstractNumId;
        }

        // Create a num instance referencing this abstractNum
        var numId = _nextNumId++;
        _numInstances.Add((numId, abstractNumId));
        return numId;
    }

    public void setVBAProject(byte[] vbaProjectData)
    {
        Guard.ThrowIfNull(vbaProjectData, nameof(vbaProjectData));
        _vbaProjectBin = vbaProjectData.ToArray();
        _vbaDataXml = null;
        _vbaProjectBinRels = null;
    }

    public void setVBAProject(Stream vbaProjectStream)
    {
        Guard.ThrowIfNull(vbaProjectStream, nameof(vbaProjectStream));
        using var ms = new MemoryStream();
        vbaProjectStream.CopyTo(ms);
        setVBAProject(ms.ToArray());
    }

    internal XWPFPictureData AddPictureData(byte[] data, int format)
    {
        var existing = _pictures.FirstOrDefault(p => p.Format == format && p.Data.SequenceEqual(data));
        if (existing is not null) return existing;
        var picture = new XWPFPictureData(data, format, _pictures.Count + 1);
        _pictures.Add(picture);
        return picture;
    }

    internal long ReserveDrawingId() => _nextDrawingId++;

    private void Load(Stream stream)
    {
        _paragraphs.Clear();
        _comments.Clear();
        _tables.Clear();
        _bodyChildren.Clear();
        _pictures.Clear();
        _nextDrawingId = 1;
        _vbaProjectBin = null;
        _vbaDataXml = null;
        _vbaProjectBinRels = null;
        _preservedEntries.Clear();
        _preservedRawBodyElements.Clear();
        _preservedRawSectPrChildren.Clear();
        _preservedBodySectPr.Clear();
        _preservedContentTypeDefaults.Clear();
        _preservedContentTypeOverrides.Clear();
        _preservedDocumentRelationships.Clear();
        _commentsModified = false;

        using var archive = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: true);

        ReadPreservedContentTypes(archive);
        ReadPreservedDocumentRelationships(archive);
        ReadPictures(archive);
        ReadComments(archive);
        ReadDocument(archive);
        ReadVbaFiles(archive);
        ReadHeadersFooters(archive);
        ReadStyles(archive);
        CollectPreservedEntries(archive);
    }

    private HashSet<string> GetModelEntryNames()
    {
        var known = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "[Content_Types].xml",
            "_rels/.rels",
            "word/document.xml",
            "word/_rels/document.xml.rels",
            "word/settings.xml",
            "word/styles.xml",
        };
        if (_numInstances.Count > 0)
            known.Add("word/numbering.xml");
        if (_commentsModified)
            known.Add("word/comments.xml");
        if (_vbaProjectBin != null)
        {
            known.Add("word/vbaProject.bin");
            if (_vbaDataXml != null) known.Add("word/vbaData.xml");
            if (_vbaProjectBinRels != null) known.Add("word/_rels/vbaProject.bin.rels");
        }
        foreach (var pic in _pictures)
            known.Add($"word/media/{pic.getFileName()}");
        return known;
    }

    private void CollectPreservedEntries(ZipArchive archive)
    {
        _preservedEntries.Clear();
        var known = GetModelEntryNames();
        foreach (var entry in archive.Entries)
        {
            var name = entry.FullName.Replace('\\', '/');
            if (known.Contains(name)) continue;
            using var s = entry.Open();
            using var ms = new MemoryStream();
            s.CopyTo(ms);
            _preservedEntries[name] = ms.ToArray();
        }
    }

    private void ReadPreservedContentTypes(ZipArchive archive)
    {
        var entry = archive.GetEntry("[Content_Types].xml");
        if (entry is null) return;

        const string ctNs = "http://schemas.openxmlformats.org/package/2006/content-types";
        using var stream = entry.Open();
        using var reader = XmlReader.Create(stream, new XmlReaderSettings { IgnoreWhitespace = false });
        while (reader.Read())
        {
            if (reader.NodeType != XmlNodeType.Element) continue;
            if (!string.Equals(reader.NamespaceURI, ctNs, StringComparison.Ordinal)) continue;

            if (reader.LocalName == "Default")
            {
                var extension = reader.GetAttribute("Extension");
                var contentType = reader.GetAttribute("ContentType");
                if (extension is null || extension.Length == 0 || contentType is null || contentType.Length == 0)
                    continue;
                if (extension.Equals("rels", StringComparison.OrdinalIgnoreCase)
                    || extension.Equals("xml", StringComparison.OrdinalIgnoreCase))
                    continue;
                _preservedContentTypeDefaults.Add(new PreservedContentTypeDefault(extension, contentType));
            }
            else if (reader.LocalName == "Override")
            {
                var partName = reader.GetAttribute("PartName");
                var contentType = reader.GetAttribute("ContentType");
                if (partName is null || partName.Length == 0 || contentType is null || contentType.Length == 0)
                    continue;
                if (IsModeledContentTypeOverride(partName))
                    continue;
                _preservedContentTypeOverrides.Add(new PreservedContentTypeOverride(partName, contentType));
            }
        }
    }

    private static bool IsModeledContentTypeOverride(string partName)
    {
        if (partName.Equals("/word/document.xml", StringComparison.OrdinalIgnoreCase)
            || partName.Equals("/word/settings.xml", StringComparison.OrdinalIgnoreCase)
            || partName.Equals("/word/numbering.xml", StringComparison.OrdinalIgnoreCase)
            || partName.Equals("/word/styles.xml", StringComparison.OrdinalIgnoreCase)
            || partName.Equals("/word/vbaProject.bin", StringComparison.OrdinalIgnoreCase)
            || partName.Equals("/word/vbaData.xml", StringComparison.OrdinalIgnoreCase))
            return true;

        var normalized = partName.TrimStart('/');
        return normalized.StartsWith("word/header", StringComparison.OrdinalIgnoreCase)
            || normalized.StartsWith("word/footer", StringComparison.OrdinalIgnoreCase)
            || normalized.StartsWith("word/media/", StringComparison.OrdinalIgnoreCase);
    }

    private void ReadPreservedDocumentRelationships(ZipArchive archive)
    {
        var entry = archive.GetEntry("word/_rels/document.xml.rels");
        if (entry is null) return;

        const string relsNs = "http://schemas.openxmlformats.org/package/2006/relationships";
        using var stream = entry.Open();
        using var reader = XmlReader.Create(stream, new XmlReaderSettings { IgnoreWhitespace = false });
        while (reader.Read())
        {
            if (reader.NodeType != XmlNodeType.Element || reader.LocalName != "Relationship")
                continue;
            if (!string.Equals(reader.NamespaceURI, relsNs, StringComparison.Ordinal))
                continue;

            var id = reader.GetAttribute("Id");
            var type = reader.GetAttribute("Type");
            var target = reader.GetAttribute("Target");
            if (id is null || id.Length == 0 || type is null || type.Length == 0
                || target is null || target.Length == 0)
                continue;
            if (IsModeledDocumentRelationship(type))
                continue;
            _preservedDocumentRelationships.Add(new PreservedRelationship(
                id,
                type,
                target,
                reader.GetAttribute("TargetMode")));
        }
    }

    private static bool IsModeledDocumentRelationship(string type) =>
        type is RelTypeSettings
            or RelTypeStyles
            or RelTypeNumbering
            or RelTypeHyperlink
            or RelTypeHeader
            or RelTypeFooter
            or RelTypeVbaProject
            or "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";

    private void ReadStyles(ZipArchive archive)
    {
        var entry = archive.GetEntry("word/styles.xml");
        if (entry is not null)
        {
            using var stream = entry.Open();
            _styles = new XWPFStyles();
            _styles.ReadStyles(stream);
        }
    }

    private void CollectHyperlinks()
    {
        _hyperlinkUrls.Clear();
        foreach (var para in _paragraphs)
        {
            CollectHyperlinksFromPara(para);
        }
        foreach (var table in _tables)
        {
            foreach (var row in table.Rows)
            {
                foreach (var cell in row.Cells)
                {
                    foreach (var para in cell.Paragraphs)
                    {
                        CollectHyperlinksFromPara(para);
                    }
                }
            }
        }
    }

    private void CollectHyperlinksFromPara(XWPFParagraph para)
    {
        foreach (var run in para.Runs)
        {
            var url = run.HyperlinkUrl;
            if (url is not null && !_hyperlinkUrls.Contains(url))
            {
                _hyperlinkUrls.Add(url);
            }
        }
    }

    private void ReadVbaFiles(ZipArchive archive)
    {
        _vbaProjectBin    = ReadEntryBytes(archive, "word/vbaProject.bin");
        _vbaDataXml       = ReadEntryBytes(archive, "word/vbaData.xml");
        _vbaProjectBinRels = ReadEntryBytes(archive, "word/_rels/vbaProject.bin.rels");
    }

    private static byte[]? ReadEntryBytes(ZipArchive archive, string path)
    {
        var entry = archive.GetEntry(path);
        if (entry is null) return null;
        using var s = entry.Open();
        using var ms = new MemoryStream();
        s.CopyTo(ms);
        return ms.Length > 0 ? ms.ToArray() : null;
    }

    private void ReadPictures(ZipArchive archive)
    {
        foreach (var entry in archive.Entries
            .Where(e => e.FullName.StartsWith("word/media/image", StringComparison.Ordinal))
            .OrderBy(e => e.FullName, StringComparer.Ordinal))
        {
            using var entryStream = entry.Open();
            using var ms = new MemoryStream();
            entryStream.CopyTo(ms);
            var ext = Path.GetExtension(entry.FullName).TrimStart('.').ToLowerInvariant();
            var format = ExtensionToFormat(ext);
            _pictures.Add(new XWPFPictureData(ms.ToArray(), format, _pictures.Count + 1));
        }
    }

    private void ReadComments(ZipArchive archive)
    {
        _comments.Clear();

        var entry = archive.GetEntry("word/comments.xml");
        if (entry is null) return;

        using var stream = entry.Open();
        using var reader = XmlReader.Create(stream, new XmlReaderSettings { IgnoreWhitespace = false });
        try
        {
            while (reader.Read())
            {
                if (reader.NodeType != XmlNodeType.Element || reader.LocalName != "comment"
                    || !string.Equals(reader.NamespaceURI, NsW, StringComparison.Ordinal))
                    continue;
                _comments.Add(ParseCommentXml(reader, MarkCommentsModified));
            }
        }
        catch (XmlException) { }
    }

    private static XWPFComment ParseCommentXml(XmlReader reader, Action? onChanged)
    {
        var id = reader.GetAttribute("id", NsW) ?? "-1";
        var author = reader.GetAttribute("author", NsW);
        var initials = reader.GetAttribute("initials", NsW);
        var date = reader.GetAttribute("date", NsW);

        var paragraphTexts = new List<string>();
        var paragraphText = new StringBuilder();
        bool inParagraph = false;
        bool inT = false;
        int relDepth = 0;
        bool done = false;

        if (!reader.IsEmptyElement)
        {
            while (!done && reader.Read())
            {
                if (reader.NodeType == XmlNodeType.Element)
                {
                    relDepth++;
                    bool isW = string.Equals(reader.NamespaceURI, NsW, StringComparison.Ordinal);
                    if (relDepth == 1 && isW && reader.LocalName == "p")
                    {
                        inParagraph = true;
                        paragraphText.Clear();
                    }
                    else if (inParagraph && isW && reader.LocalName == "t")
                    {
                        inT = true;
                    }
                    if (reader.IsEmptyElement) relDepth--;
                }
                else if (reader.NodeType == XmlNodeType.EndElement)
                {
                    bool isW = string.Equals(reader.NamespaceURI, NsW, StringComparison.Ordinal);
                    if (inT && isW && reader.LocalName == "t") inT = false;
                    if (inParagraph && relDepth == 1 && isW && reader.LocalName == "p")
                    {
                        paragraphTexts.Add(paragraphText.ToString());
                        inParagraph = false;
                    }
                    relDepth--;
                    if (relDepth < 0) done = true;
                }
                else if (reader.NodeType == XmlNodeType.Text && inT)
                {
                    paragraphText.Append(reader.Value);
                }
            }
        }

        return new XWPFComment(id, author, initials, date, string.Join("\n", paragraphTexts), onChanged);
    }

    private void ReadHeadersFooters(ZipArchive archive)
    {
        // Build relId → part path map from document relationships
        var hfRelMap = new Dictionary<string, string>(StringComparer.Ordinal);
        var relsEntry = archive.GetEntry("word/_rels/document.xml.rels");
        if (relsEntry is not null)
        {
            using var rs = relsEntry.Open();
            using var rr = XmlReader.Create(rs, new XmlReaderSettings { IgnoreWhitespace = false });
            while (rr.Read())
            {
                if (rr.NodeType != XmlNodeType.Element || rr.LocalName != "Relationship") continue;
                var relType = rr.GetAttribute("Type");
                var isH = relType is not null && relType.EndsWith("/relationships/header", StringComparison.Ordinal);
                var isF = relType is not null && relType.EndsWith("/relationships/footer", StringComparison.Ordinal);
                if (isH || isF)
                {
                    var id = rr.GetAttribute("Id");
                    var target = rr.GetAttribute("Target");
                    if (id is not null && target is not null)
                    {
                        var partPath = target.StartsWith("/") ? target.Substring(1) : $"word/{target}";
                        // Normalize separators
                        partPath = partPath.Replace('\\', '/');
                        hfRelMap[id] = partPath;
                    }
                }
            }
        }

        // Read header/footer XML text, routing to correct variant field by filename
        foreach (var kv in hfRelMap)
        {
            var entry = archive.GetEntry(kv.Value);
            if (entry is null) continue;
            using var s = entry.Open();
            using var hr = XmlReader.Create(s, new XmlReaderSettings { IgnoreWhitespace = false });

            // Determine variant from filename (header1=default, header2=first, header3=even)
            var fileName = Path.GetFileNameWithoutExtension(kv.Value);
            bool isHeader = fileName.StartsWith("header", StringComparison.Ordinal);
            int idx = isHeader
                ? (fileName == "header1" ? 0 : fileName == "header2" ? 1 : 2)
                : (fileName == "footer1" ? 0 : fileName == "footer2" ? 1 : 2);

            while (hr.Read())
            {
                if (hr.NodeType == XmlNodeType.Element && hr.LocalName == "t")
                {
                    var text = hr.ReadElementContentAsString();
                    if (isHeader)
                    {
                        if (idx == 0) _headerText = (_headerText ?? "") + text;
                        else if (idx == 1) _headerFirstText = (_headerFirstText ?? "") + text;
                        else _headerEvenText = (_headerEvenText ?? "") + text;
                    }
                    else
                    {
                        if (idx == 0) _footerText = (_footerText ?? "") + text;
                        else if (idx == 1) _footerFirstText = (_footerFirstText ?? "") + text;
                        else _footerEvenText = (_footerEvenText ?? "") + text;
                    }
                    continue;
                }
            }
        }
    }

    private void ReadDocument(ZipArchive archive)
    {
        // Build relId → picture-index map from document relationships
        var picIndexByRelId = BuildPictureRelMap(archive);
        var hyperlinkRelMap = BuildHyperlinkRelMap(archive);

        var docEntry = archive.GetEntry("word/document.xml");
        if (docEntry is null) return;

        using var docStream = docEntry.Open();
        using var reader = XmlReader.Create(docStream, new XmlReaderSettings { IgnoreWhitespace = false });

        XWPFParagraph? currentParagraph = null;
        XWPFRun? currentRun = null;
        XWPFTable? currentTable = null;
        XWPFTableRow? currentRow = null;
        XWPFTableCell? currentTableCell = null;
        string? hyperlinkRelId = null;
        bool inPPr = false;
        bool inRPr = false;

        // Field parsing state
        bool inField = false;
        bool fieldPastSeparate = false;
        string? fieldInstruction = null;
        string fieldResult = "";

        // Inline image parsing state
        bool inDrawing = false;
        long wpInlineExtCx = 0, wpInlineExtCy = 0;
        string? wpDocPrDescr = null;
        long wpDocPrId = 0;
        string? blipEmbed = null;
        int xfrmRot = 0;
        bool inInline = false;

        // Table property parsing state
        bool inTblPr = false;
        bool inTrPr = false;
        bool inTcPr = false;

        // Raw XML preservation state
        int bodyDepth = -1;
        int paragraphDepth = -1;
        int runDepth = -1;
        int sectPrDepth = -1;
        bool skipRead = false;

        while (skipRead || reader.Read())
        {
            skipRead = false;

            if (reader.NodeType == XmlNodeType.Element)
            {
                // Track w:body entry to know when we're inside body
                if (reader.NamespaceURI == NsW && reader.LocalName == "body")
                {
                    bodyDepth = reader.Depth;
                    continue;
                }

                // Capture unknown direct children of w:body that the model doesn't handle
                // (SDT/content controls, altChunk, bookmarks, track changes, etc.)
                if (bodyDepth >= 0 && reader.Depth == bodyDepth + 1)
                {
                    if (reader.NamespaceURI == NsW)
                    {
                        if (reader.LocalName is not ("p" or "tbl" or "sectPr" or "body"))
                        {
                            AddPreservedRawBodyElement(reader.ReadOuterXml());
                            skipRead = true;
                            continue;
                        }
                    }
                    else
                    {
                        // Non-W namespace direct body child (e.g., mc:AlternateContent)
                        AddPreservedRawBodyElement(reader.ReadOuterXml());
                        skipRead = true;
                        continue;
                    }
                }

                // Track w:p entry for inline SDT/bookmark/track-changes preservation
                if (reader.NamespaceURI == NsW && reader.LocalName == "p")
                {
                    paragraphDepth = reader.Depth;
                }

                // Capture unknown direct children of w:p (inline SDT, bookmarks, tracked changes, etc.)
                if (paragraphDepth >= 0 && reader.Depth == paragraphDepth + 1)
                {
                    bool isKnown = reader.NamespaceURI == NsW
                        && reader.LocalName is "pPr" or "r" or "hyperlink";
                    if (!isKnown && currentParagraph is not null)
                    {
                        currentParagraph.addPreservedRawElement(reader.ReadOuterXml());
                        skipRead = true;
                        continue;
                    }
                }

                // Capture body-level section breaks (mid-document sectPr between body children)
                // The ENTIRE element is captured as raw XML for verbatim re-emission.
                // The LAST body-level sectPr is also captured here; its children are later
                // parsed for model state via ParseFinalSectPr() after the while loop.
                if (reader.NamespaceURI == NsW && reader.LocalName == "sectPr"
                    && bodyDepth >= 0 && reader.Depth == bodyDepth + 1)
                {
                    _preservedBodySectPr.Add(reader.ReadOuterXml());
                    skipRead = true;
                    continue;
                }
            }

            if (reader.NodeType == XmlNodeType.Element)
            {
                switch (reader.NamespaceURI)
                {
                    case NsW:
                        switch (reader.LocalName)
                        {
                            case "p":
                                if (currentTableCell is not null)
                                    currentParagraph = currentTableCell.addParagraph();
                                else
                                    currentParagraph = createParagraph();
                                break;
                            case "hyperlink":
                                hyperlinkRelId = GetAttributeByLocalName(reader, "id");
                                break;
                            case "r" when !inRPr:
                                if (currentParagraph is not null)
                                {
                                    currentRun = currentParagraph.createRun();
                                    runDepth = reader.Depth;
                                    if (hyperlinkRelId is not null && hyperlinkRelMap.TryGetValue(hyperlinkRelId, out var url))
                                        currentRun.setHyperlink(url);
                                }
                                break;
                            case "rPr":
                                inRPr = true;
                                break;
                            case "pPr":
                                inPPr = true;
                                break;
                            case "b" when inRPr:
                                currentRun?.setBold(true);
                                break;
                            case "i" when inRPr:
                                currentRun?.setItalic(true);
                                break;
                            case "u" when inRPr:
                                currentRun?.setUnderline(true);
                                break;
                            case "strike" when inRPr:
                                currentRun?.setStrike(true);
                                break;
                            case "rFonts" when inRPr:
                                var asciiAttr = reader.GetAttribute("w:ascii");
                                if (asciiAttr is not null)
                                    currentRun?.setFontName(asciiAttr);
                                break;
                            case "sz" when inRPr:
                                var szVal = reader.GetAttribute("w:val");
                                if (szVal is not null && double.TryParse(szVal, NumberStyles.Float, CultureInfo.InvariantCulture, out var halfPt))
                                    currentRun?.setFontSize(halfPt / 2.0);
                                break;
                            case "color" when inRPr:
                                var colorVal = reader.GetAttribute("w:val");
                                if (colorVal is not null)
                                    currentRun?.setColor(colorVal);
                                break;
                            case "pStyle" when inPPr:
                                var pStyleVal = reader.GetAttribute("w:val");
                                if (pStyleVal is not null && currentParagraph is not null)
                                    currentParagraph.setStyle(pStyleVal);
                                break;
                            case "jc" when inPPr:
                                var jcVal = reader.GetAttribute("w:val");
                                if (jcVal is not null && currentParagraph is not null)
                                {
                                    currentParagraph.setAlignment(jcVal switch
                                    {
                                        "left" => ParagraphAlignment.Left,
                                        "center" => ParagraphAlignment.Center,
                                        "right" => ParagraphAlignment.Right,
                                        "both" => ParagraphAlignment.Both,
                                        _ => ParagraphAlignment.Left
                                    });
                                }
                                break;
                            case "ind" when inPPr:
                                var leftAttr = reader.GetAttribute("w:left");
                                var rightAttr = reader.GetAttribute("w:right");
                                var firstLineAttr = reader.GetAttribute("w:firstLine");
                                var hangingAttr = reader.GetAttribute("w:hanging");
                                if (leftAttr is not null && int.TryParse(leftAttr, out var l))
                                    currentParagraph?.setIndentationLeft(l);
                                if (rightAttr is not null && int.TryParse(rightAttr, out var r))
                                    currentParagraph?.setIndentationRight(r);
                                if (firstLineAttr is not null && int.TryParse(firstLineAttr, out var fl))
                                    currentParagraph?.setIndentationFirstLine(fl);
                                if (hangingAttr is not null && int.TryParse(hangingAttr, out var h))
                                    currentParagraph?.setIndentationHanging(h);
                                break;
                            case "spacing" when inPPr:
                                var beforeAttr = reader.GetAttribute("w:before");
                                var afterAttr = reader.GetAttribute("w:after");
                                var lineAttr = reader.GetAttribute("w:line");
                                var lineRuleAttr = reader.GetAttribute("w:lineRule");
                                if (beforeAttr is not null && int.TryParse(beforeAttr, out var b))
                                    currentParagraph?.setSpacingBefore(b);
                                if (afterAttr is not null && int.TryParse(afterAttr, out var a))
                                    currentParagraph?.setSpacingAfter(a);
                                if (lineAttr is not null && int.TryParse(lineAttr, out var ln))
                                    currentParagraph?.setSpacingBetween(ln);
                                if (lineRuleAttr is not null)
                                    currentParagraph?.setLineSpacingRule(lineRuleAttr switch
                                    {
                                        "atLeast" => LineSpacingRule.AtLeast,
                                        "exact" => LineSpacingRule.Exact,
                                        _ => LineSpacingRule.Auto
                                    });
                                break;
                            case "numPr" when inPPr:
                                break;
                            case "numId" when inPPr:
                                var numIdAttr = reader.GetAttribute("w:val");
                                if (numIdAttr is not null && int.TryParse(numIdAttr, out var nid))
                                    currentParagraph?.setNumId(nid);
                                break;
                            case "ilvl" when inPPr:
                                var ilvlAttr = reader.GetAttribute("w:val");
                                if (ilvlAttr is not null && int.TryParse(ilvlAttr, out var iv))
                                    currentParagraph?.setIlvl(iv);
                                break;
                            case "pgSz":
                                var pgW = reader.GetAttribute("w:w");
                                var pgH = reader.GetAttribute("w:h");
                                var orient = reader.GetAttribute("w:orient");
                                if (pgW is not null && long.TryParse(pgW, NumberStyles.Integer, CultureInfo.InvariantCulture, out var pw))
                                    _pageWidth = pw;
                                if (pgH is not null && long.TryParse(pgH, NumberStyles.Integer, CultureInfo.InvariantCulture, out var ph))
                                    _pageHeight = ph;
                                if (orient == "landscape")
                                    _pageOrientation = "landscape";
                                break;
                            case "pgMar":
                            {
                                var tp = reader.GetAttribute("w:top");
                                var rp = reader.GetAttribute("w:right");
                                var bp = reader.GetAttribute("w:bottom");
                                var lp = reader.GetAttribute("w:left");
                                var hdp = reader.GetAttribute("w:header");
                                var ftp = reader.GetAttribute("w:footer");
                                if (tp is not null && long.TryParse(tp, out var mt)) _marginTop = mt;
                                if (rp is not null && long.TryParse(rp, out var mr)) _marginRight = mr;
                                if (bp is not null && long.TryParse(bp, out var mb)) _marginBottom = mb;
                                if (lp is not null && long.TryParse(lp, out var ml)) _marginLeft = ml;
                                if (hdp is not null && long.TryParse(hdp, out var mh)) _marginHeader = mh;
                                if (ftp is not null && long.TryParse(ftp, out var mf)) _marginFooter = mf;
                                break;
                            }
                            case "cols":
                            {
                                var colNum = reader.GetAttribute("w:num");
                                var colSpace = reader.GetAttribute("w:space");
                                if (colNum is not null && int.TryParse(colNum, out var cn) && cn > 1)
                                    _columnCount = cn;
                                if (colSpace is not null && long.TryParse(colSpace, out var cs))
                                    _columnSpacing = cs;
                                break;
                            }
                            case "sectPr" when inPPr && currentParagraph is not null:
                            {
                                // Paragraph-level section break — capture raw XML for re-emission
                                currentParagraph.setPreservedSectPr(reader.ReadOuterXml());
                                skipRead = true;
                                continue;
                            }
                            case "t" when currentRun is not null && !inRPr:
                                var tContent = reader.ReadElementContentAsString();
                                currentRun.setText(tContent);
                                if (inField && fieldPastSeparate)
                                {
                                    fieldResult += tContent;
                                }
                                continue;
                            case "txbxContent" when currentRun is not null:
                                foreach (var textBoxText in ExtractTextBoxParagraphTexts(reader.ReadOuterXml()))
                                {
                                    currentRun.addTextBoxText(textBoxText);
                                }
                                skipRead = true;
                                continue;
                            case "fldChar":
                                var fldCharType = reader.GetAttribute("w:fldCharType");
                                if (fldCharType == "begin" && currentParagraph is not null)
                                {
                                    inField = true;
                                    fieldPastSeparate = false;
                                    fieldInstruction = null;
                                    fieldResult = "";
                                }
                                else if (fldCharType == "separate" && inField)
                                {
                                    fieldPastSeparate = true;
                                }
                                else if (fldCharType == "end" && inField && currentParagraph is not null)
                                {
                                    currentParagraph.addField(fieldInstruction ?? "", fieldResult);
                                    inField = false;
                                    fieldPastSeparate = false;
                                    fieldInstruction = null;
                                    fieldResult = "";
                                }
                                break;
                            case "instrText" when inField:
                                var instrText = reader.ReadElementContentAsString();
                                fieldInstruction = instrText;
                                continue;
                            case "drawing" when currentRun is not null:
                                inDrawing = true;
                                wpInlineExtCx = wpInlineExtCy = 0;
                                wpDocPrDescr = null;
                                wpDocPrId = 0;
                                blipEmbed = null;
                                xfrmRot = 0;
                                inInline = false;
                                break;
                            case "tbl":
                                currentTable = createTable();
                                break;
                            case "tr" when currentTable is not null:
                                currentRow = currentTable.createRow();
                                break;
                            case "tc" when currentRow is not null:
                                currentTableCell = currentRow.createCell();
                                break;

                            // Table property context tracking
                            case "tblPr" when currentTable is not null:
                                inTblPr = true;
                                break;
                            case "trPr" when currentRow is not null:
                                inTrPr = true;
                                break;
                            case "tcPr" when currentTableCell is not null:
                                inTcPr = true;
                                break;

                            // tblPr modeled children
                            case "tblW" when inTblPr && currentTable is not null:
                                if (long.TryParse(reader.GetAttribute("w:w"), out var tw))
                                    currentTable.setWidth(tw, reader.GetAttribute("w:type") ?? "dxa");
                                break;
                            case "tblStyle" when inTblPr && currentTable is not null:
                                currentTable.setTableStyle(reader.GetAttribute("w:val"));
                                break;
                            case "tblBorders" when inTblPr && currentTable is not null:
                            {
                                ParseTableBorders(reader.ReadOuterXml(), currentTable);
                                skipRead = true;
                                continue;
                            }

                            // trPr modeled children
                            case "trHeight" when inTrPr && currentRow is not null:
                            {
                                var hAttr = reader.GetAttribute("w:val");
                                var rule = reader.GetAttribute("w:hRule");
                                if (hAttr is not null && long.TryParse(hAttr, out var th))
                                    currentRow.setHeight(th, rule ?? "atLeast");
                                break;
                            }
                            case "tblHeader" when inTrPr && currentRow is not null:
                                currentRow.setHeader(true);
                                break;

                            // tcPr modeled children
                            case "tcW" when inTcPr && currentTableCell is not null:
                                if (long.TryParse(reader.GetAttribute("w:w"), out var tcw))
                                    currentTableCell.setWidth(tcw, reader.GetAttribute("w:type") ?? "dxa");
                                break;
                            case "gridSpan" when inTcPr && currentTableCell is not null:
                                if (int.TryParse(reader.GetAttribute("w:val"), out var gs) && gs > 1)
                                    currentTableCell.setGridSpan(gs);
                                break;
                            case "vMerge" when inTcPr && currentTableCell is not null:
                            {
                                var v = reader.GetAttribute("w:val");
                                currentTableCell.setVMerge(v ?? "continue");
                                break;
                            }
                            case "hMerge" when inTcPr && currentTableCell is not null:
                            {
                                var v = reader.GetAttribute("w:val");
                                currentTableCell.setHMerge(v ?? "continue");
                                break;
                            }
                            case "vAlign" when inTcPr && currentTableCell is not null:
                                currentTableCell.setVAlign(reader.GetAttribute("w:val"));
                                break;
                        }
                        break;

                    case NsWp:
                        if (reader.LocalName == "inline") inInline = true;
                        if (reader.LocalName == "extent" && inInline)
                        {
                            if (long.TryParse(reader.GetAttribute("cx"), out var cx)) wpInlineExtCx = cx;
                            if (long.TryParse(reader.GetAttribute("cy"), out var cy)) wpInlineExtCy = cy;
                        }
                        if (reader.LocalName == "docPr" && inInline)
                        {
                            wpDocPrDescr = reader.GetAttribute("descr");
                            if (long.TryParse(reader.GetAttribute("id"), out var id)) wpDocPrId = id;
                        }
                        // Anchored (floating) images: capture entire anchor as raw XML
                        if (reader.LocalName == "anchor" && currentRun is not null)
                        {
                            var anchorXml = reader.ReadOuterXml();
                            skipRead = true;

                            foreach (var textBoxText in ExtractTextBoxParagraphTexts(anchorXml))
                            {
                                currentRun.addTextBoxText(textBoxText);
                            }

                            // Parse anchor XML to register picture data
                            var (embed, cx, cy, descr, rot) = ParseAnchorBlip(anchorXml);
                            if (embed is not null
                                && picIndexByRelId.TryGetValue(embed, out var picIdx)
                                && (uint)picIdx < (uint)_pictures.Count)
                            {
                                var picData = _pictures[picIdx];
                                var drawingId = ReserveDrawingId();
                                var picture = new XWPFPicture(picData, descr ?? picData.getFileName(),
                                    cx, cy, embed, drawingId);
                                picture.SetRotationAttribute(rot);
                                AttachPictureToRun(currentRun, picture);
                            }

                            // Store raw XML for re-emission on write
                            currentRun.addRawAnchorXml(anchorXml);
                            continue;
                        }
                        break;

                    case NsA:
                        if (reader.LocalName == "blip")
                        {
                            blipEmbed = GetAttributeByLocalName(reader, "embed");
                        }
                        if (reader.LocalName == "xfrm" && inInline)
                        {
                            if (int.TryParse(reader.GetAttribute("rot"), out var rot)) xfrmRot = rot;
                        }
                        break;
                }

                // Capture unmodeled direct children of body-level w:sectPr
                // (pgBorders, lnNumType, docGrid, formProt, etc.) that fell through
                // the main switch because they aren't pgSz/pgMar/cols.
                if (sectPrDepth >= 0 && reader.Depth == sectPrDepth + 1 && reader.NodeType == XmlNodeType.Element)
                {
                    _preservedRawSectPrChildren.Add(reader.ReadOuterXml());
                    skipRead = true;
                    continue;
                }

                // Capture unmodeled children of tblPr, trPr, tcPr as raw XML for re-emission
                // NOTE: The property container elements (tblPr, trPr, tcPr) themselves MUST be
                // excluded so that their EndElement handler can reset the inXxxPr flag.
                if (inTblPr && currentTable is not null && reader.NodeType == XmlNodeType.Element
                    && reader.LocalName is not ("tblPr" or "tblW" or "tblStyle" or "tblBorders"))
                {
                    currentTable.addPreservedRawTblPrChild(reader.ReadOuterXml());
                    skipRead = true;
                    continue;
                }
                if (inTrPr && currentRow is not null && reader.NodeType == XmlNodeType.Element
                    && reader.LocalName is not ("trPr" or "trHeight" or "tblHeader"))
                {
                    currentRow.addPreservedRawTrPrChild(reader.ReadOuterXml());
                    skipRead = true;
                    continue;
                }
                if (inTcPr && currentTableCell is not null && reader.NodeType == XmlNodeType.Element
                    && reader.LocalName is not ("tcPr" or "tcW" or "gridSpan" or "vMerge" or "hMerge" or "vAlign"))
                {
                    currentTableCell.addPreservedRawTcPrChild(reader.ReadOuterXml());
                    skipRead = true;
                    continue;
                }
                // Preserve unmodeled direct children of w:r, including w:commentReference.
                if (!inRPr && currentRun is not null && runDepth >= 0
                    && reader.NodeType == XmlNodeType.Element
                    && reader.Depth == runDepth + 1
                    && reader.LocalName is not ("rPr" or "t" or "drawing" or "fldChar" or "instrText"))
                {
                    currentRun.addPreservedRawContentElement(reader.ReadOuterXml());
                    skipRead = true;
                    continue;
                }
                // Capture unmodeled children of run properties (rPr) as raw XML for re-emission
                // Run-level rPr (inside w:r)
                if (inRPr && currentRun is not null && reader.NodeType == XmlNodeType.Element
                    && reader.LocalName is not ("rPr" or "b" or "i" or "u" or "strike" or "rFonts" or "sz" or "color"))
                {
                    currentRun.addPreservedRawRPrChild(reader.ReadOuterXml());
                    skipRead = true;
                    continue;
                }
                // Paragraph-level rPr (inside w:pPr)
                if (inRPr && inPPr && currentParagraph is not null && currentRun is null && reader.NodeType == XmlNodeType.Element
                    && reader.LocalName is not ("rPr" or "b" or "i" or "u" or "strike" or "rFonts" or "sz" or "color"))
                {
                    currentParagraph.addPreservedRawPPrRPrChild(reader.ReadOuterXml());
                    skipRead = true;
                    continue;
                }
            }
            else if (reader.NodeType == XmlNodeType.EndElement)
            {
                switch (reader.NamespaceURI)
                {
                    case NsW:
                        switch (reader.LocalName)
                        {
                            case "p":
                                currentParagraph = null;
                                currentRun = null;
                                inPPr = false;
                                inRPr = false;
                                paragraphDepth = -1;
                                runDepth = -1;
                                break;
                            case "hyperlink":
                                hyperlinkRelId = null;
                                break;
                            case "r" when !inRPr:
                                currentRun = null;
                                runDepth = -1;
                                break;
                            case "rPr":
                                inRPr = false;
                                break;
                            case "pPr":
                                inPPr = false;
                                break;
                            case "drawing" when inDrawing:
                                inDrawing = false;
                                inInline = false;
                                break;
                            case "tblPr":
                                inTblPr = false;
                                break;
                            case "trPr":
                                inTrPr = false;
                                break;
                            case "tcPr":
                                inTcPr = false;
                                break;
                            case "tbl":
                                currentTable = null;
                                break;
                            case "tr":
                                currentRow = null;
                                break;
                            case "tc":
                                currentTableCell = null;
                                break;
                            case "sectPr":
                                sectPrDepth = -1;
                                break;
                            case "body":
                                bodyDepth = -1;
                                break;
                        }
                        break;
                    case NsWp:
                        if (reader.LocalName == "inline" && inDrawing && inInline
                            && blipEmbed is not null
                            && currentRun is not null)
                        {
                            if (picIndexByRelId.TryGetValue(blipEmbed, out var picIdx)
                                && (uint)picIdx < (uint)_pictures.Count)
                            {
                                var picData = _pictures[picIdx];
                                var drawingId = ReserveDrawingId();
                                var picture = new XWPFPicture(picData, wpDocPrDescr ?? picData.getFileName(),
                                    wpInlineExtCx, wpInlineExtCy, blipEmbed, drawingId);
                                picture.SetRotationAttribute(xfrmRot);
                                // Attach picture to run via internal list
                                AttachPictureToRun(currentRun, picture);
                            }
                            inInline = false;
                            blipEmbed = null;
                        }
                        break;
                }
            }
        }

        // After the read loop, parse the LAST body-level sectPr for model state
        // (pgSz, pgMar, cols, and unmodeled children like pgBorders, lnNumType, etc.)
        ParseFinalSectPr();
    }

    private void ParseFinalSectPr()
    {
        if (_preservedBodySectPr.Count == 0)
            return;

        var lastSectPr = _preservedBodySectPr[_preservedBodySectPr.Count - 1];
        using var sr = new StringReader(lastSectPr);
        using var reader = XmlReader.Create(sr, new XmlReaderSettings { IgnoreWhitespace = false });

        _preservedRawSectPrChildren.Clear();

        bool skipRead = false;
        while (skipRead || reader.Read())
        {
            skipRead = false;
            if (reader.NodeType != XmlNodeType.Element) continue;

            switch (reader.LocalName)
            {
                case "sectPr":
                    // The outer element — skip to process children
                    break;
                case "pgSz":
                {
                    var pgW = reader.GetAttribute("w:w");
                    var pgH = reader.GetAttribute("w:h");
                    var orient = reader.GetAttribute("w:orient");
                    if (pgW is not null && long.TryParse(pgW, NumberStyles.Integer, CultureInfo.InvariantCulture, out var pw))
                        _pageWidth = pw;
                    if (pgH is not null && long.TryParse(pgH, NumberStyles.Integer, CultureInfo.InvariantCulture, out var ph))
                        _pageHeight = ph;
                    if (orient == "landscape")
                        _pageOrientation = "landscape";
                    break;
                }
                case "pgMar":
                {
                    var tp = reader.GetAttribute("w:top");
                    var rp = reader.GetAttribute("w:right");
                    var bp = reader.GetAttribute("w:bottom");
                    var lp = reader.GetAttribute("w:left");
                    var hdp = reader.GetAttribute("w:header");
                    var ftp = reader.GetAttribute("w:footer");
                    if (tp is not null && long.TryParse(tp, out var mt)) _marginTop = mt;
                    if (rp is not null && long.TryParse(rp, out var mr)) _marginRight = mr;
                    if (bp is not null && long.TryParse(bp, out var mb)) _marginBottom = mb;
                    if (lp is not null && long.TryParse(lp, out var ml)) _marginLeft = ml;
                    if (hdp is not null && long.TryParse(hdp, out var mh)) _marginHeader = mh;
                    if (ftp is not null && long.TryParse(ftp, out var mf)) _marginFooter = mf;
                    break;
                }
                case "cols":
                {
                    var colNum = reader.GetAttribute("w:num");
                    var colSpace = reader.GetAttribute("w:space");
                    if (colNum is not null && int.TryParse(colNum, out var cn) && cn > 1)
                        _columnCount = cn;
                    if (colSpace is not null && long.TryParse(colSpace, out var cs))
                        _columnSpacing = cs;
                    break;
                }
                case "headerReference":
                case "footerReference":
                    // These are handled by model — skip raw preservation
                    break;
                default:
                    // Unmodeled children (pgBorders, lnNumType, docGrid, formProt, etc.)
                    // captured as raw for re-emission in model's final sectPr
                    if (reader.IsEmptyElement)
                    {
                        // Self-closing: build XML manually to avoid .NET 8 ReadOuterXml bug
                        // Do NOT advance reader here — the while loop handles that via reader.Read()
                        _preservedRawSectPrChildren.Add(CaptureEmptyElementXml(reader));
                    }
                    else
                    {
                        // Non-empty element: ReadOuterXml advances past it, so use skipRead
                        _preservedRawSectPrChildren.Add(reader.ReadOuterXml());
                        skipRead = true;
                    }
                    break;
            }
        }
    }

    /// <summary>
    /// Captures the current XmlReader self-closing element as a raw XML string
    /// WITHOUT advancing the reader. The caller/loop handles advancement.
    /// Does NOT use ReadOuterXml() for self-closing elements to work around a .NET 8
    /// bug where ReadOuterXml() on a self-closing element with no trailing whitespace
    /// before the next sibling silently consumes that sibling.
    /// Only call for self-closing (IsEmptyElement == true) elements.
    /// </summary>
    private static string CaptureEmptyElementXml(XmlReader reader)
    {
        // Self-closing element: build XML manually
        var sb = new StringBuilder();
        sb.Append('<');
        if (!string.IsNullOrEmpty(reader.Prefix))
            sb.Append(reader.Prefix).Append(':');
        sb.Append(reader.LocalName);

        if (reader.HasAttributes)
        {
            for (int i = 0; i < reader.AttributeCount; i++)
            {
                reader.MoveToAttribute(i);
                sb.Append(' ').Append(reader.Name).Append("=\"").Append(reader.Value).Append('"');
            }
            reader.MoveToElement();
        }

        sb.Append(" />");
        // Do NOT advance reader — the calling loop's reader.Read() handles that
        return sb.ToString();
    }

    private Dictionary<string, int> BuildPictureRelMap(ZipArchive archive)
    {
        var map = new Dictionary<string, int>(StringComparer.Ordinal);
        var relsEntry = archive.GetEntry("word/_rels/document.xml.rels");
        if (relsEntry is null) return map;

        using var relsStream = relsEntry.Open();
        using var reader = XmlReader.Create(relsStream, new XmlReaderSettings { IgnoreWhitespace = false });
        while (reader.Read())
        {
            if (reader.NodeType != XmlNodeType.Element || reader.LocalName != "Relationship") continue;
            if (!string.Equals(reader.GetAttribute("Type"),
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
                StringComparison.Ordinal)) continue;
            var relId = reader.GetAttribute("Id");
            var target = reader.GetAttribute("Target"); // e.g. "media/image1.jpeg"
            if (relId is null || target is null) continue;
            var mediaFile = Path.GetFileNameWithoutExtension(target); // "image1"
            if (mediaFile.StartsWith("image", StringComparison.OrdinalIgnoreCase)
                && int.TryParse(mediaFile.Substring("image".Length), out var oneBasedIndex))
            {
                map[relId] = oneBasedIndex - 1;
            }
        }
        return map;
    }

    private Dictionary<string, string> BuildHyperlinkRelMap(ZipArchive archive)
    {
        var map = new Dictionary<string, string>(StringComparer.Ordinal);
        var relsEntry = archive.GetEntry("word/_rels/document.xml.rels");
        if (relsEntry is null) return map;

        using var relsStream = relsEntry.Open();
        using var reader = XmlReader.Create(relsStream, new XmlReaderSettings { IgnoreWhitespace = false });
        while (reader.Read())
        {
            if (reader.NodeType != XmlNodeType.Element || reader.LocalName != "Relationship") continue;
            if (!string.Equals(reader.GetAttribute("Type"),
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
                StringComparison.Ordinal)) continue;
            var relId = reader.GetAttribute("Id");
            var target = reader.GetAttribute("Target");
            if (relId is not null && target is not null)
                map[relId] = target;
        }
        return map;
    }

    private static void AttachPictureToRun(XWPFRun run, XWPFPicture picture)
    {
        run.AttachPicture(picture);
    }

    private static string? GetAttributeByLocalName(XmlReader reader, string localName)
    {
        if (!reader.HasAttributes) return null;
        while (reader.MoveToNextAttribute())
        {
            if (reader.LocalName == localName)
            {
                var value = reader.Value;
                reader.MoveToElement();
                return value;
            }
        }
        reader.MoveToElement();
        return null;
    }

    private static void ParseTableBorders(string bordersXml, XWPFTable table)
    {
        using var sr = new StringReader(bordersXml);
        using var reader = XmlReader.Create(sr, new XmlReaderSettings { IgnoreWhitespace = false });

        while (reader.Read())
        {
            if (reader.NodeType != XmlNodeType.Element)
                continue;

            var position = reader.LocalName switch
            {
                "top" => TableBorderPosition.Top,
                "bottom" => TableBorderPosition.Bottom,
                "left" => TableBorderPosition.Left,
                "right" => TableBorderPosition.Right,
                "insideH" => TableBorderPosition.InsideH,
                "insideV" => TableBorderPosition.InsideV,
                _ => (TableBorderPosition?)null
            };

            if (position is null)
                continue;

            var type = ParseBorderXmlValue(GetAttributeByLocalName(reader, "val"));
            var size = ParseIntAttribute(reader, "sz", -1);
            var space = ParseIntAttribute(reader, "space", -1);
            var color = GetAttributeByLocalName(reader, "color") ?? "auto";
            table.SetBorder(position.Value, type, size, space, color);
        }
    }

    private static int ParseIntAttribute(XmlReader reader, string localName, int defaultValue)
    {
        var value = GetAttributeByLocalName(reader, localName);
        return int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var parsed)
            ? parsed
            : defaultValue;
    }

    private static XWPFTable.XWPFBorderType ParseBorderXmlValue(string? value)
    {
        return value switch
        {
            "nil" => XWPFTable.XWPFBorderType.Nil,
            "none" => XWPFTable.XWPFBorderType.None,
            "single" => XWPFTable.XWPFBorderType.Single,
            "thick" => XWPFTable.XWPFBorderType.Thick,
            "double" => XWPFTable.XWPFBorderType.Double,
            "dotted" => XWPFTable.XWPFBorderType.Dotted,
            "dashed" => XWPFTable.XWPFBorderType.Dashed,
            "dotDash" => XWPFTable.XWPFBorderType.DotDash,
            "dotDotDash" => XWPFTable.XWPFBorderType.DotDotDash,
            "triple" => XWPFTable.XWPFBorderType.Triple,
            "thinThickSmallGap" => XWPFTable.XWPFBorderType.ThinThickSmallGap,
            "thickThinSmallGap" => XWPFTable.XWPFBorderType.ThickThinSmallGap,
            "thinThickThinSmallGap" => XWPFTable.XWPFBorderType.ThinThickThinSmallGap,
            "thinThickMediumGap" => XWPFTable.XWPFBorderType.ThinThickMediumGap,
            "thickThinMediumGap" => XWPFTable.XWPFBorderType.ThickThinMediumGap,
            "thinThickThinMediumGap" => XWPFTable.XWPFBorderType.ThinThickThinMediumGap,
            "thinThickLargeGap" => XWPFTable.XWPFBorderType.ThinThickLargeGap,
            "thickThinLargeGap" => XWPFTable.XWPFBorderType.ThickThinLargeGap,
            "thinThickThinLargeGap" => XWPFTable.XWPFBorderType.ThinThickThinLargeGap,
            "wave" => XWPFTable.XWPFBorderType.Wave,
            "doubleWave" => XWPFTable.XWPFBorderType.DoubleWave,
            "dashSmallGap" => XWPFTable.XWPFBorderType.DashSmallGap,
            "dashDotStroked" => XWPFTable.XWPFBorderType.DashDotStroked,
            "threeDEmboss" => XWPFTable.XWPFBorderType.ThreeDEmboss,
            "threeDEngrave" => XWPFTable.XWPFBorderType.ThreeDEngrave,
            "outset" => XWPFTable.XWPFBorderType.Outset,
            "inset" => XWPFTable.XWPFBorderType.Inset,
            _ => XWPFTable.XWPFBorderType.None
        };
    }

    private static string ToBorderXmlValue(XWPFTable.XWPFBorderType type)
    {
        return type switch
        {
            XWPFTable.XWPFBorderType.Nil => "nil",
            XWPFTable.XWPFBorderType.None => "none",
            XWPFTable.XWPFBorderType.Single => "single",
            XWPFTable.XWPFBorderType.Thick => "thick",
            XWPFTable.XWPFBorderType.Double => "double",
            XWPFTable.XWPFBorderType.Dotted => "dotted",
            XWPFTable.XWPFBorderType.Dashed => "dashed",
            XWPFTable.XWPFBorderType.DotDash => "dotDash",
            XWPFTable.XWPFBorderType.DotDotDash => "dotDotDash",
            XWPFTable.XWPFBorderType.Triple => "triple",
            XWPFTable.XWPFBorderType.ThinThickSmallGap => "thinThickSmallGap",
            XWPFTable.XWPFBorderType.ThickThinSmallGap => "thickThinSmallGap",
            XWPFTable.XWPFBorderType.ThinThickThinSmallGap => "thinThickThinSmallGap",
            XWPFTable.XWPFBorderType.ThinThickMediumGap => "thinThickMediumGap",
            XWPFTable.XWPFBorderType.ThickThinMediumGap => "thickThinMediumGap",
            XWPFTable.XWPFBorderType.ThinThickThinMediumGap => "thinThickThinMediumGap",
            XWPFTable.XWPFBorderType.ThinThickLargeGap => "thinThickLargeGap",
            XWPFTable.XWPFBorderType.ThickThinLargeGap => "thickThinLargeGap",
            XWPFTable.XWPFBorderType.ThinThickThinLargeGap => "thinThickThinLargeGap",
            XWPFTable.XWPFBorderType.Wave => "wave",
            XWPFTable.XWPFBorderType.DoubleWave => "doubleWave",
            XWPFTable.XWPFBorderType.DashSmallGap => "dashSmallGap",
            XWPFTable.XWPFBorderType.DashDotStroked => "dashDotStroked",
            XWPFTable.XWPFBorderType.ThreeDEmboss => "threeDEmboss",
            XWPFTable.XWPFBorderType.ThreeDEngrave => "threeDEngrave",
            XWPFTable.XWPFBorderType.Outset => "outset",
            XWPFTable.XWPFBorderType.Inset => "inset",
            _ => "none"
        };
    }

    /// <summary>
    /// Parses a captured wp:anchor XML string to extract picture-related attributes:
    /// blip embed (relationship ID), extent cx/cy, docPr descr, and xfrm rot.
    /// </summary>
    private static (string? embed, long cx, long cy, string? descr, int rot) ParseAnchorBlip(string anchorXml)
    {
        string? embed = null, descr = null;
        long cx = 0, cy = 0;
        int rot = 0;

        using var sr = new StringReader(anchorXml);
        using var reader = XmlReader.Create(sr, new XmlReaderSettings { IgnoreWhitespace = false });

        while (reader.Read())
        {
            if (reader.NodeType != XmlNodeType.Element) continue;
            switch (reader.LocalName)
            {
                case "blip":
                    embed = GetAttributeByLocalName(reader, "embed");
                    break;
                case "extent":
                    if (long.TryParse(reader.GetAttribute("cx"), out var ecx)) cx = ecx;
                    if (long.TryParse(reader.GetAttribute("cy"), out var ecy)) cy = ecy;
                    break;
                case "docPr":
                    descr = reader.GetAttribute("descr");
                    break;
                case "xfrm":
                    if (int.TryParse(reader.GetAttribute("rot"), out var xrot)) rot = xrot;
                    break;
            }
        }

        return (embed, cx, cy, descr, rot);
    }

    private static IReadOnlyList<string> ExtractTextBoxParagraphTexts(string xml)
    {
        var paragraphs = new List<string>();
        using var sr = new StringReader(xml);
        using var reader = XmlReader.Create(sr, new XmlReaderSettings { IgnoreWhitespace = false });

        int textBoxDepth = -1;
        int paragraphDepth = -1;
        StringBuilder? paragraphText = null;

        while (reader.Read())
        {
            if (reader.NodeType == XmlNodeType.Element)
            {
                if (reader.NamespaceURI == NsW && reader.LocalName == "txbxContent")
                {
                    textBoxDepth = reader.Depth;
                    continue;
                }

                if (textBoxDepth >= 0 && reader.NamespaceURI == NsW && reader.LocalName == "p")
                {
                    paragraphDepth = reader.Depth;
                    paragraphText = new StringBuilder();
                    continue;
                }

                if (paragraphText is not null && reader.NamespaceURI == NsW)
                {
                    switch (reader.LocalName)
                    {
                        case "t":
                            paragraphText.Append(reader.ReadElementContentAsString());
                            break;
                        case "tab":
                            paragraphText.Append('\t');
                            break;
                        case "br":
                        case "cr":
                            paragraphText.Append('\n');
                            break;
                    }
                }
            }
            else if (reader.NodeType == XmlNodeType.EndElement)
            {
                if (reader.NamespaceURI == NsW && reader.LocalName == "p" && reader.Depth == paragraphDepth)
                {
                    paragraphs.Add(paragraphText?.ToString() ?? string.Empty);
                    paragraphText = null;
                    paragraphDepth = -1;
                }
                else if (reader.NamespaceURI == NsW && reader.LocalName == "txbxContent" && reader.Depth == textBoxDepth)
                {
                    textBoxDepth = -1;
                }
            }
        }

        return paragraphs;
    }

    private static int ExtensionToFormat(string ext) => ext switch
    {
        "jpg" or "jpeg" => XWPFPictureData.PICTURE_TYPE_JPEG,
        "png" => XWPFPictureData.PICTURE_TYPE_PNG,
        "gif" => XWPFPictureData.PICTURE_TYPE_GIF,
        "dib" => XWPFPictureData.PICTURE_TYPE_DIB,
        "tif" or "tiff" => XWPFPictureData.PICTURE_TYPE_TIFF,
        "eps" => XWPFPictureData.PICTURE_TYPE_EPS,
        "bmp" => XWPFPictureData.PICTURE_TYPE_BMP,
        "wpg" => XWPFPictureData.PICTURE_TYPE_WPG,
        "emf" => XWPFPictureData.PICTURE_TYPE_EMF,
        "wmf" => XWPFPictureData.PICTURE_TYPE_WMF,
        "pict" => XWPFPictureData.PICTURE_TYPE_PICT,
        _ => throw new NotSupportedException($"Image extension '{ext}' is not supported.")
    };

    private void WriteContentTypes(PoiXmlWriter writer)
    {
        writer.WriteStartElement("Types");
        writer.WriteAttributeString("xmlns", "http://schemas.openxmlformats.org/package/2006/content-types");
        WriteDefault(writer, "rels", "application/vnd.openxmlformats-package.relationships+xml");
        foreach (var ext in _pictures
            .Select(p => p.Extension)
            .Distinct(StringComparer.Ordinal)
            .OrderBy(e => e, StringComparer.Ordinal))
        {
            WriteDefault(writer, ext, _pictures.First(p => p.Extension == ext).ContentType);
        }
        var writtenDefaultExtensions = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "rels",
            "xml"
        };
        foreach (var picture in _pictures)
            writtenDefaultExtensions.Add(picture.Extension);
        foreach (var preservedDefault in _preservedContentTypeDefaults)
        {
            if (!writtenDefaultExtensions.Add(preservedDefault.Extension))
                continue;
            WriteDefault(writer, preservedDefault.Extension, preservedDefault.ContentType);
        }
        WriteDefault(writer, "xml", "application/xml");
        // docm uses macroEnabled content type; docx uses the standard one.
        WriteOverride(writer, "/word/document.xml",
            _vbaProjectBin != null ? ContentTypeDocm : ContentTypeDocx);
        WriteOverride(writer, "/word/settings.xml", ContentTypeSettings);
        if (_numInstances.Count > 0)
            WriteOverride(writer, "/word/numbering.xml", ContentTypeNumbering);
        WriteOverride(writer, "/word/styles.xml", ContentTypeStyles);
        {
            int hIdx = 1;
            foreach (var hText in new[] { _headerText, _headerFirstText, _headerEvenText })
                if (!string.IsNullOrEmpty(hText))
                    WriteOverride(writer, $"/word/header{hIdx++}.xml", ContentTypeHeader);
        }
        {
            int fIdx = 1;
            foreach (var fText in new[] { _footerText, _footerFirstText, _footerEvenText })
                if (!string.IsNullOrEmpty(fText))
                    WriteOverride(writer, $"/word/footer{fIdx++}.xml", ContentTypeFooter);
        }
        if (_vbaProjectBin != null)
        {
            WriteOverride(writer, "/word/vbaProject.bin", ContentTypeVbaProject);
            if (_vbaDataXml != null)
                WriteOverride(writer, "/word/vbaData.xml", ContentTypeVbaData);
        }
        var writtenOverridePartNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "/word/document.xml",
            "/word/settings.xml",
            "/word/styles.xml"
        };
        if (_numInstances.Count > 0)
            writtenOverridePartNames.Add("/word/numbering.xml");
        if (_commentsModified)
        {
            WriteOverride(writer, "/word/comments.xml", ContentTypeComments);
            writtenOverridePartNames.Add("/word/comments.xml");
        }
        foreach (var preservedOverride in _preservedContentTypeOverrides)
        {
            if (!writtenOverridePartNames.Add(preservedOverride.PartName))
                continue;
            WriteOverride(writer, preservedOverride.PartName, preservedOverride.ContentType);
        }
        writer.WriteEndElement();
    }

    private static void WriteRootRelationships(PoiXmlWriter writer)
    {
        writer.WriteStartElement("Relationships");
        writer.WriteAttributeString("xmlns", "http://schemas.openxmlformats.org/package/2006/relationships");
        WriteRelationship(writer, "rId1", "word/document.xml",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument");
        writer.WriteEndElement();
    }

    private void WriteDocumentRelationships(PoiXmlWriter writer)
    {
        writer.WriteStartElement("Relationships");
        writer.WriteAttributeString("xmlns", "http://schemas.openxmlformats.org/package/2006/relationships");
        // rId1: settings
        WriteRelationship(writer, "rId1", "settings.xml", RelTypeSettings);
        // rId2: styles
        WriteRelationship(writer, "rId2", "styles.xml", RelTypeStyles);
        // images: rId{pic.Index + ImageRelIdOffset}
        foreach (var pic in _pictures)
        {
            WriteRelationship(writer, $"rId{pic.Index + ImageRelIdOffset}", $"media/{pic.getFileName()}",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image");
        }
        // hyperlinks: after images, before numbering/vba
        for (int i = 0; i < _hyperlinkUrls.Count; i++)
        {
            var relId = $"rId{_pictures.Count + 3 + i}";
            WriteRelationship(writer, relId, _hyperlinkUrls[i], RelTypeHyperlink, "External");
        }
        int offset = _hyperlinkUrls.Count;
        // numbering
        if (_numInstances.Count > 0)
        {
            var numRelId = $"rId{_pictures.Count + 3 + offset}";
            WriteRelationship(writer, numRelId, "numbering.xml", RelTypeNumbering);
            offset++;
        }
        // vbaProject
        if (_vbaProjectBin != null)
        {
            var vbaRelId = $"rId{_pictures.Count + 3 + offset}";
            WriteRelationship(writer, vbaRelId, "vbaProject.bin", RelTypeVbaProject);
            offset++;
        }
        // headers (variants: default, first, even)
        {
            int hFileIdx = 1;
            foreach (var hText in new[] { _headerText, _headerFirstText, _headerEvenText })
            {
                if (string.IsNullOrEmpty(hText)) continue;
                offset++;
                var relId = $"rId{_pictures.Count + 3 + offset - 1}";
                WriteRelationship(writer, relId, $"header{hFileIdx++}.xml", RelTypeHeader);
            }
        }
        // footers (variants: default, first, even)
        {
            int fFileIdx = 1;
            foreach (var fText in new[] { _footerText, _footerFirstText, _footerEvenText })
            {
                if (string.IsNullOrEmpty(fText)) continue;
                offset++;
                var relId = $"rId{_pictures.Count + 3 + offset - 1}";
                WriteRelationship(writer, relId, $"footer{fFileIdx++}.xml", RelTypeFooter);
            }
        }
        if (_commentsModified)
        {
            var commentsRelId = GetPreservedCommentsRelationshipId()
                ?? GetNextDocumentRelationshipId(_pictures.Count + 3 + offset);
            WriteRelationship(writer, commentsRelId, "comments.xml", RelTypeComments);
        }
        foreach (var relationship in _preservedDocumentRelationships)
        {
            if (_commentsModified && relationship.Type.Equals(RelTypeComments, StringComparison.Ordinal))
                continue;
            if (relationship.TargetMode is null)
                WriteRelationship(writer, relationship.Id, relationship.Target, relationship.Type);
            else
                WriteRelationship(writer, relationship.Id, relationship.Target, relationship.Type, relationship.TargetMode);
        }
        writer.WriteEndElement();
    }

    private string? GetPreservedCommentsRelationshipId() =>
        _preservedDocumentRelationships.FirstOrDefault(r => r.Type.Equals(RelTypeComments, StringComparison.Ordinal))?.Id;

    private string GetNextDocumentRelationshipId(int start)
    {
        var used = new HashSet<string>(StringComparer.Ordinal);
        used.Add("rId1");
        used.Add("rId2");
        foreach (var picture in _pictures)
            used.Add($"rId{picture.Index + ImageRelIdOffset}");
        for (int i = 0; i < _hyperlinkUrls.Count; i++)
            used.Add($"rId{_pictures.Count + 3 + i}");
        foreach (var relationship in _preservedDocumentRelationships)
            used.Add(relationship.Id);

        var index = Math.Max(3, start);
        while (used.Contains($"rId{index}"))
            index++;
        return $"rId{index}";
    }

    private void WriteComments(PoiXmlWriter writer)
    {
        writer.WriteStartElement("w", "comments");
        writer.WriteAttributeString("xmlns:w", NsW);
        foreach (var comment in _comments.OrderBy(c => int.TryParse(c.getId(), out var id) ? id : int.MaxValue))
        {
            writer.WriteStartElement("w", "comment");
            writer.WriteAttributeString("w", "id", comment.getId());
            if (comment.getAuthor() is not null)
                writer.WriteAttributeString("w", "author", comment.getAuthor()!);
            if (comment.getInitials() is not null)
                writer.WriteAttributeString("w", "initials", comment.getInitials()!);
            if (comment.getDate() is not null)
                writer.WriteAttributeString("w", "date", comment.getDate()!);
            foreach (var paragraphText in SplitCommentParagraphs(comment.getText()))
            {
                writer.WriteStartElement("w", "p");
                writer.WriteStartElement("w", "r");
                writer.WriteStartElement("w", "t");
                writer.WriteAttributeString("xml:space", "preserve");
                writer.WriteString(paragraphText);
                writer.WriteEndElement();
                writer.WriteEndElement();
                writer.WriteEndElement();
            }
            writer.WriteEndElement();
        }
        writer.WriteEndElement();
    }

    private static IEnumerable<string> SplitCommentParagraphs(string text)
    {
        var normalized = text.Replace("\r\n", "\n").Replace('\r', '\n');
        var paragraphs = normalized.Split('\n');
        return paragraphs.Length == 0 ? new[] { string.Empty } : paragraphs;
    }

    private static void WriteSettings(PoiXmlWriter writer)
    {
        writer.WriteStartElement("w", "settings");
        writer.WriteAttributeString("xmlns:w", NsW);
        writer.WriteEndElement();
    }

    private static void WriteHeader(PoiXmlWriter writer, string text)
    {
        writer.WriteStartElement("w", "hdr");
        writer.WriteAttributeString("xmlns:w", NsW);
        writer.WriteStartElement("w", "p");
        writer.WriteStartElement("w", "r");
        writer.WriteStartElement("w", "t");
        writer.WriteString(text);
        writer.WriteEndElement();
        writer.WriteEndElement();
        writer.WriteEndElement();
        writer.WriteEndElement();
    }

    private static void WriteFooter(PoiXmlWriter writer, string text)
    {
        writer.WriteStartElement("w", "ftr");
        writer.WriteAttributeString("xmlns:w", NsW);
        writer.WriteStartElement("w", "p");
        writer.WriteStartElement("w", "r");
        writer.WriteStartElement("w", "t");
        writer.WriteString(text);
        writer.WriteEndElement();
        writer.WriteEndElement();
        writer.WriteEndElement();
        writer.WriteEndElement();
    }

    private void WriteStyles(PoiXmlWriter writer)
    {
        XWPFStyles.WriteDefaultStyles(writer);
    }

    private void WriteDocument(PoiXmlWriter writer)
    {
        writer.WriteStartElement("w", "document");
        writer.WriteAttributeString("xmlns:w", NsW);
        writer.WriteAttributeString("xmlns:r", NsR);
        if (_pictures.Count > 0)
        {
            writer.WriteAttributeString("xmlns:wp", NsWp);
            writer.WriteAttributeString("xmlns:a", NsA);
            writer.WriteAttributeString("xmlns:pic", NsPic);
        }
        writer.WriteStartElement("w", "body");
        foreach (var child in _bodyChildren)
        {
            WriteBodyChild(writer, child);
        }
        // Emit mid-document body-level section breaks (all except the final one)
        // The final body-level sectPr is emitted as the model's sectPr below.
        for (int i = 0; i < _preservedBodySectPr.Count - 1; i++)
        {
            writer.WriteRaw(_preservedBodySectPr[i]);
        }
        writer.WriteStartElement("w", "sectPr");
        // pgSz: page size + orientation
        writer.WriteStartElement("w", "pgSz");
        writer.WriteAttributeString("w", "w", _pageWidth.ToString(CultureInfo.InvariantCulture));
        writer.WriteAttributeString("w", "h", _pageHeight.ToString(CultureInfo.InvariantCulture));
        if (_pageOrientation is not null)
            writer.WriteAttributeString("w", "orient", _pageOrientation);
        writer.WriteEndElement();
        // pgMar: margins
        writer.WriteStartElement("w", "pgMar");
        writer.WriteAttributeString("w", "top", _marginTop.ToString(CultureInfo.InvariantCulture));
        writer.WriteAttributeString("w", "right", _marginRight.ToString(CultureInfo.InvariantCulture));
        writer.WriteAttributeString("w", "bottom", _marginBottom.ToString(CultureInfo.InvariantCulture));
        writer.WriteAttributeString("w", "left", _marginLeft.ToString(CultureInfo.InvariantCulture));
        writer.WriteAttributeString("w", "header", _marginHeader.ToString(CultureInfo.InvariantCulture));
        writer.WriteAttributeString("w", "footer", _marginFooter.ToString(CultureInfo.InvariantCulture));
        writer.WriteEndElement();
        // cols: newspaper columns
        if (_columnCount > 1 || _columnSpacing > 0)
        {
            writer.WriteStartElement("w", "cols");
            writer.WriteAttributeString("w", "num", _columnCount.ToString(CultureInfo.InvariantCulture));
            if (_columnSpacing > 0)
                writer.WriteAttributeString("w", "space", _columnSpacing.ToString(CultureInfo.InvariantCulture));
            writer.WriteEndElement();
        }
        // Re-emit preserved raw sectPr children (pgBorders, lnNumType, docGrid, formProt, etc.)
        foreach (var raw in _preservedRawSectPrChildren)
        {
            writer.WriteRaw(raw);
        }
        // headerReference / footerReference (variants: default, first, even)
        {
            int relOffset = 3 + _pictures.Count + _hyperlinkUrls.Count;
            if (_numInstances.Count > 0) relOffset++;
            if (_vbaProjectBin != null) relOffset++;

            // Headers in type-order: default, first, even
            var headerVariants = new (string? text, string type)[]
            {
                (_headerText, "default"),
                (_headerFirstText, "first"),
                (_headerEvenText, "even"),
            };
            foreach (var (text, type) in headerVariants)
            {
                if (string.IsNullOrEmpty(text)) continue;
                writer.WriteStartElement("w", "headerReference");
                writer.WriteAttributeString("w", "type", type);
                writer.WriteAttributeString("r", "id", $"rId{relOffset++}");
                writer.WriteEndElement();
            }

            // Footers in type-order: default, first, even
            var footerVariants = new (string? text, string type)[]
            {
                (_footerText, "default"),
                (_footerFirstText, "first"),
                (_footerEvenText, "even"),
            };
            foreach (var (text, type) in footerVariants)
            {
                if (string.IsNullOrEmpty(text)) continue;
                writer.WriteStartElement("w", "footerReference");
                writer.WriteAttributeString("w", "type", type);
                writer.WriteAttributeString("r", "id", $"rId{relOffset++}");
                writer.WriteEndElement();
            }
        }
        writer.WriteEndElement(); // sectPr
        writer.WriteEndElement(); // body
        writer.WriteEndElement(); // document
    }

    private void WriteBodyChild(PoiXmlWriter writer, BodyChild child)
    {
        switch (child.Kind)
        {
            case BodyChildKind.Paragraph when child.Paragraph is not null:
                WriteParagraph(writer, child.Paragraph);
                break;
            case BodyChildKind.Table when child.Table is not null:
                WriteTable(writer, child.Table);
                break;
            case BodyChildKind.Raw when child.RawXml is not null:
                writer.WriteRaw(child.RawXml);
                break;
        }
    }

    private void WriteParagraph(PoiXmlWriter writer, XWPFParagraph para)
    {
        writer.WriteStartElement("w", "p");

        if (para.Alignment is not null || para.IndentLeft != 0
            || para.IndentRight != 0 || para.IndentFirstLine != 0 || para.IndentHanging != 0
            || para.SpacingBefore != 0 || para.SpacingAfter != 0 || para.SpacingBetween != 0
            || para.NumId is not null
            || para.PreservedSectPr is not null
            || para.HasPreservedRawPPrRPrChildren
            || para.getStyleID() is not null)
        {
            writer.WriteStartElement("w", "pPr");
            var styleId = para.getStyleID();
            if (styleId is not null)
            {
                writer.WriteStartElement("w", "pStyle");
                writer.WriteAttributeString("w", "val", styleId);
                writer.WriteEndElement();
            }
            if (para.Alignment is not null)
            {
                writer.WriteStartElement("w", "jc");
                writer.WriteAttributeString("w", "val", para.Alignment switch
                {
                    ParagraphAlignment.Left => "left",
                    ParagraphAlignment.Center => "center",
                    ParagraphAlignment.Right => "right",
                    ParagraphAlignment.Both => "both",
                    _ => "left"
                });
                writer.WriteEndElement();
            }
            if (para.IndentLeft != 0 || para.IndentRight != 0
                || para.IndentFirstLine != 0 || para.IndentHanging != 0)
            {
                writer.WriteStartElement("w", "ind");
                if (para.IndentLeft != 0)
                    writer.WriteAttributeString("w", "left", para.IndentLeft.ToString(CultureInfo.InvariantCulture));
                if (para.IndentRight != 0)
                    writer.WriteAttributeString("w", "right", para.IndentRight.ToString(CultureInfo.InvariantCulture));
                if (para.IndentFirstLine != 0)
                    writer.WriteAttributeString("w", "firstLine", para.IndentFirstLine.ToString(CultureInfo.InvariantCulture));
                if (para.IndentHanging != 0)
                    writer.WriteAttributeString("w", "hanging", para.IndentHanging.ToString(CultureInfo.InvariantCulture));
                writer.WriteEndElement();
            }
            if (para.SpacingBefore != 0 || para.SpacingAfter != 0
                || para.SpacingBetween != 0)
            {
                writer.WriteStartElement("w", "spacing");
                if (para.SpacingBefore != 0)
                    writer.WriteAttributeString("w", "before", para.SpacingBefore.ToString(CultureInfo.InvariantCulture));
                if (para.SpacingAfter != 0)
                    writer.WriteAttributeString("w", "after", para.SpacingAfter.ToString(CultureInfo.InvariantCulture));
                if (para.SpacingBetween != 0)
                {
                    writer.WriteAttributeString("w", "line", para.SpacingBetween.ToString(CultureInfo.InvariantCulture));
                    writer.WriteAttributeString("w", "lineRule", para.LineRule switch
                    {
                        LineSpacingRule.AtLeast => "atLeast",
                        LineSpacingRule.Exact => "exact",
                        _ => "auto"
                    });
                }
                writer.WriteEndElement();
            }
            if (para.NumId is not null)
            {
                writer.WriteStartElement("w", "numPr");
                writer.WriteStartElement("w", "ilvl");
                writer.WriteAttributeString("w", "val", para.Ilvl.ToString(CultureInfo.InvariantCulture));
                writer.WriteEndElement();
                writer.WriteStartElement("w", "numId");
                writer.WriteAttributeString("w", "val", para.NumId.Value.ToString(CultureInfo.InvariantCulture));
                writer.WriteEndElement();
                writer.WriteEndElement();
            }
            // Re-emit paragraph-level section properties (sectPr for section break)
            if (para.PreservedSectPr is not null)
            {
                writer.WriteRaw(para.PreservedSectPr);
            }
            // Re-emit preserved raw pPr/rPr children (e.g., w:shd at paragraph level)
            foreach (var rawChild in para.PreservedRawPPrRPrChildren)
            {
                writer.WriteRaw(rawChild);
            }
            writer.WriteEndElement(); // pPr
        }

        foreach (var child in para.Children)
        {
            switch (child.Kind)
            {
                case XWPFParagraph.ChildKind.Run when child.Run is not null:
                    WriteRun(writer, child.Run);
                    break;
                case XWPFParagraph.ChildKind.Field when child.Field is not null:
                    WriteField(writer, child.Field);
                    break;
                case XWPFParagraph.ChildKind.Raw when child.RawXml is not null:
                    writer.WriteRaw(child.RawXml);
                    break;
            }
        }

        writer.WriteEndElement();
    }

    private void WriteField(PoiXmlWriter writer, XWPFField field)
    {
        // begin
        writer.WriteStartElement("w", "r");
        writer.WriteStartElement("w", "fldChar");
        writer.WriteAttributeString("w", "fldCharType", "begin");
        writer.WriteEndElement();
        writer.WriteEndElement();

        // instrText
        writer.WriteStartElement("w", "r");
        writer.WriteStartElement("w", "instrText");
        writer.WriteAttributeString("xml:space", "preserve");
        writer.WriteString(field.Instruction);
        writer.WriteEndElement();
        writer.WriteEndElement();

        // separate
        writer.WriteStartElement("w", "r");
        writer.WriteStartElement("w", "fldChar");
        writer.WriteAttributeString("w", "fldCharType", "separate");
        writer.WriteEndElement();
        writer.WriteEndElement();

        // result (if non-empty)
        if (!string.IsNullOrEmpty(field.Result))
        {
            writer.WriteStartElement("w", "r");
            writer.WriteStartElement("w", "t");
            writer.WriteString(field.Result);
            writer.WriteEndElement();
            writer.WriteEndElement();
        }

        // end
        writer.WriteStartElement("w", "r");
        writer.WriteStartElement("w", "fldChar");
        writer.WriteAttributeString("w", "fldCharType", "end");
        writer.WriteEndElement();
        writer.WriteEndElement();
    }

    private void WriteTable(PoiXmlWriter writer, XWPFTable table)
    {
        writer.WriteStartElement("w", "tbl");

        // tblPr: modeled properties + preserved raw children
        if (table.getWidth() > 0 || table.getWidthType() is not null
            || table.getTableStyle() is not null
            || table.Borders.Count > 0
            || table.PreservedRawTblPrChildren.Count > 0)
        {
            writer.WriteStartElement("w", "tblPr");
            var tableStyle = table.getTableStyle();
            if (tableStyle is not null)
            {
                writer.WriteStartElement("w", "tblStyle");
                writer.WriteAttributeString("w", "val", tableStyle);
                writer.WriteEndElement();
            }
            if (table.getWidth() > 0 || table.getWidthType() is not null)
            {
                writer.WriteStartElement("w", "tblW");
                writer.WriteAttributeString("w", "w", table.getWidth().ToString(CultureInfo.InvariantCulture));
                writer.WriteAttributeString("w", "type", table.getWidthType() ?? "dxa");
                writer.WriteEndElement();
            }
            if (table.Borders.Count > 0)
                WriteTableBorders(writer, table);
            foreach (var raw in table.PreservedRawTblPrChildren)
                writer.WriteRaw(raw);
            writer.WriteEndElement(); // tblPr
        }
        else
        {
            // Always emit at least a basic tblPr for validity
            writer.WriteStartElement("w", "tblPr");
            writer.WriteStartElement("w", "tblW");
            writer.WriteAttributeString("w", "w", "5000");
            writer.WriteAttributeString("w", "type", "dxa");
            writer.WriteEndElement();
            writer.WriteEndElement(); // tblPr
        }

        if (table.GridColWidths.Count > 0)
        {
            writer.WriteStartElement("w", "tblGrid");
            foreach (var w in table.GridColWidths)
            {
                writer.WriteStartElement("w", "gridCol");
                writer.WriteAttributeString("w", "w", w.ToString(CultureInfo.InvariantCulture));
                writer.WriteEndElement();
            }
            writer.WriteEndElement(); // tblGrid
        }
        foreach (var row in table.Rows)
        {
            writer.WriteStartElement("w", "tr");

            // trPr
            if (row.getHeight() > 0 || row.isHeader() || row.PreservedRawTrPrChildren.Count > 0)
            {
                writer.WriteStartElement("w", "trPr");
                if (row.isHeader())
                {
                    writer.WriteStartElement("w", "tblHeader");
                    writer.WriteEndElement();
                }
                if (row.getHeight() > 0)
                {
                    writer.WriteStartElement("w", "trHeight");
                    writer.WriteAttributeString("w", "val", row.getHeight().ToString(CultureInfo.InvariantCulture));
                    var heightRule = row.getHeightRule();
                    if (heightRule is not null)
                        writer.WriteAttributeString("w", "hRule", heightRule);
                    writer.WriteEndElement();
                }
                foreach (var raw in row.PreservedRawTrPrChildren)
                    writer.WriteRaw(raw);
                writer.WriteEndElement(); // trPr
            }

            foreach (var cell in row.Cells)
            {
                writer.WriteStartElement("w", "tc");

                // tcPr
                if (cell.getWidth() > 0 || cell.getWidthType() is not null
                    || cell.getGridSpan() > 1 || cell.getVMerge() is not null
                    || cell.getHMerge() is not null || cell.getVAlign() is not null
                    || cell.PreservedRawTcPrChildren.Count > 0)
                {
                    writer.WriteStartElement("w", "tcPr");
                    if (cell.getWidth() > 0 || cell.getWidthType() is not null)
                    {
                        writer.WriteStartElement("w", "tcW");
                        writer.WriteAttributeString("w", "w", cell.getWidth().ToString(CultureInfo.InvariantCulture));
                        writer.WriteAttributeString("w", "type", cell.getWidthType() ?? "dxa");
                        writer.WriteEndElement();
                    }
                    if (cell.getGridSpan() > 1)
                    {
                        writer.WriteStartElement("w", "gridSpan");
                        writer.WriteAttributeString("w", "val", cell.getGridSpan().ToString(CultureInfo.InvariantCulture));
                        writer.WriteEndElement();
                    }
                    var vMerge = cell.getVMerge();
                    if (vMerge is not null)
                    {
                        writer.WriteStartElement("w", "vMerge");
                        writer.WriteAttributeString("w", "val", vMerge);
                        writer.WriteEndElement();
                    }
                    var hMerge = cell.getHMerge();
                    if (hMerge is not null)
                    {
                        writer.WriteStartElement("w", "hMerge");
                        writer.WriteAttributeString("w", "val", hMerge);
                        writer.WriteEndElement();
                    }
                    if (cell.getVAlign() is not null)
                    {
                        writer.WriteStartElement("w", "vAlign");
                        writer.WriteAttributeString("w", "val", cell.getVAlign()!);
                        writer.WriteEndElement();
                    }
                    foreach (var raw in cell.PreservedRawTcPrChildren)
                        writer.WriteRaw(raw);
                    writer.WriteEndElement(); // tcPr
                }

                foreach (var para in cell.Paragraphs)
                {
                    WriteParagraph(writer, para);
                }
                writer.WriteEndElement(); // tc
            }
            writer.WriteEndElement(); // tr
        }
        writer.WriteEndElement(); // tbl
    }

    private static void WriteTableBorders(PoiXmlWriter writer, XWPFTable table)
    {
        writer.WriteStartElement("w", "tblBorders");
        WriteTableBorder(writer, table, TableBorderPosition.Top, "top");
        WriteTableBorder(writer, table, TableBorderPosition.Left, "left");
        WriteTableBorder(writer, table, TableBorderPosition.Bottom, "bottom");
        WriteTableBorder(writer, table, TableBorderPosition.Right, "right");
        WriteTableBorder(writer, table, TableBorderPosition.InsideH, "insideH");
        WriteTableBorder(writer, table, TableBorderPosition.InsideV, "insideV");
        writer.WriteEndElement();
    }

    private static void WriteTableBorder(PoiXmlWriter writer, XWPFTable table, TableBorderPosition position, string elementName)
    {
        if (!table.Borders.TryGetValue(position, out var border))
            return;

        writer.WriteStartElement("w", elementName);
        writer.WriteAttributeString("w", "val", ToBorderXmlValue(border.Type));
        writer.WriteAttributeString("w", "sz", border.Size.ToString(CultureInfo.InvariantCulture));
        writer.WriteAttributeString("w", "space", border.Space.ToString(CultureInfo.InvariantCulture));
        writer.WriteAttributeString("w", "color", border.Color);
        writer.WriteEndElement();
    }

    private void WriteRun(PoiXmlWriter writer, XWPFRun run)
    {
        if (run.HasContent)
        {
            bool isHyperlink = run.HyperlinkUrl is not null;
            if (isHyperlink)
            {
                // Ensure relId is assigned
                if (run.HyperlinkRelId is null)
                {
                    int hyperlinkIndex = _hyperlinkUrls.IndexOf(run.HyperlinkUrl!);
                    run.HyperlinkRelId = hyperlinkIndex >= 0
                        ? $"rId{_pictures.Count + 3 + hyperlinkIndex}"
                        : throw new InvalidOperationException("Hyperlink URL not registered.");
                }
                writer.WriteStartElement("w", "hyperlink");
                writer.WriteAttributeString("r", "id", run.HyperlinkRelId);
                writer.WriteAttributeString("w", "history", "1");
            }

            writer.WriteStartElement("w", "r");
            if (run.Bold || run.Italic || run.Underline || run.Strike
                || run.FontName is not null || run.FontSize > 0 || run.Color is not null
                || run.HasPreservedRawRPrChildren)
            {
                writer.WriteStartElement("w", "rPr");
                if (run.Bold)
                {
                    writer.WriteStartElement("w", "b");
                    writer.WriteAttributeString("w", "val", "on");
                    writer.WriteEndElement();
                }
                if (run.Italic)
                {
                    writer.WriteStartElement("w", "i");
                    writer.WriteAttributeString("w", "val", "on");
                    writer.WriteEndElement();
                }
                if (run.Underline)
                {
                    writer.WriteStartElement("w", "u");
                    writer.WriteAttributeString("w", "val", "single");
                    writer.WriteEndElement();
                }
                if (run.Strike)
                {
                    writer.WriteStartElement("w", "strike");
                    writer.WriteAttributeString("w", "val", "on");
                    writer.WriteEndElement();
                }
                if (run.FontName is not null)
                {
                    writer.WriteStartElement("w", "rFonts");
                    writer.WriteAttributeString("w", "ascii", run.FontName);
                    writer.WriteAttributeString("w", "hAnsi", run.FontName);
                    writer.WriteEndElement();
                }
                if (run.FontSize > 0)
                {
                    writer.WriteStartElement("w", "sz");
                    writer.WriteAttributeString("w", "val", ((int)(run.FontSize * 2)).ToString(CultureInfo.InvariantCulture));
                    writer.WriteEndElement();
                }
                if (run.Color is not null)
                {
                    writer.WriteStartElement("w", "color");
                    writer.WriteAttributeString("w", "val", run.Color);
                    writer.WriteEndElement();
                }
                // Re-emit preserved raw rPr children (e.g., w:shd, w:highlight, etc.)
                foreach (var rawChild in run.PreservedRawRPrChildren)
                {
                    writer.WriteRaw(rawChild);
                }
                writer.WriteEndElement(); // rPr
            }
            foreach (var child in run.ContentChildren)
            {
                switch (child.Kind)
                {
                    case XWPFRun.ContentChildKind.Text when run.TextValue is not null:
                        writer.WriteStartElement("w", "t");
                        writer.WriteAttributeString("xml:space", "preserve");
                        writer.WriteString(run.TextValue);
                        writer.WriteEndElement(); // t
                        break;
                    case XWPFRun.ContentChildKind.Raw when child.RawXml is not null:
                        writer.WriteRaw(child.RawXml);
                        break;
                }
            }
            writer.WriteEndElement(); // r

            if (isHyperlink)
            {
                writer.WriteEndElement(); // hyperlink
            }
        }
        foreach (var picture in run.Pictures)
        {
            WriteInlinePicture(writer, picture);
        }
        // Re-emit anchored (floating) images from raw XML preservation
        foreach (var rawAnchor in run.RawAnchorXml)
        {
            writer.WriteStartElement("w", "r");
            writer.WriteStartElement("w", "drawing");
            writer.WriteRaw(rawAnchor);
            writer.WriteEndElement(); // w:drawing
            writer.WriteEndElement(); // w:r
        }
    }

    private static void WriteInlinePicture(PoiXmlWriter writer, XWPFPicture picture)
    {
        var cx = picture.Width.ToString(CultureInfo.InvariantCulture);
        var cy = picture.Height.ToString(CultureInfo.InvariantCulture);
        var id = picture.DrawingId.ToString(CultureInfo.InvariantCulture);
        var name = $"Drawing {id}";
        var picName = $"Picture {id}";
        var descr = picture.Filename;

        writer.WriteStartElement("w", "r");
        writer.WriteStartElement("w", "drawing");

        writer.WriteStartElement("wp", "inline");
        writer.WriteAttributeString("distT", "0");
        writer.WriteAttributeString("distR", "0");
        writer.WriteAttributeString("distB", "0");
        writer.WriteAttributeString("distL", "0");

        writer.WriteStartElement("wp", "extent");
        writer.WriteAttributeString("cx", cx);
        writer.WriteAttributeString("cy", cy);
        writer.WriteEndElement();

        writer.WriteStartElement("wp", "effectExtent");
        writer.WriteAttributeString("l", "0");
        writer.WriteAttributeString("t", "0");
        writer.WriteAttributeString("r", "0");
        writer.WriteAttributeString("b", "0");
        writer.WriteEndElement();

        writer.WriteStartElement("wp", "docPr");
        writer.WriteAttributeString("id", id);
        writer.WriteAttributeString("name", name);
        writer.WriteAttributeString("descr", descr);
        writer.WriteEndElement();

        writer.WriteStartElement("wp", "cNvGraphicFramePr");
        writer.WriteStartElement("a", "graphicFrameLocks");
        writer.WriteAttributeString("noChangeAspect", "1");
        writer.WriteEndElement();
        writer.WriteEndElement();

        writer.WriteStartElement("a", "graphic");
        writer.WriteStartElement("a", "graphicData");
        writer.WriteAttributeString("uri", NsPic);

        writer.WriteStartElement("pic", "pic");

        writer.WriteStartElement("pic", "nvPicPr");
        writer.WriteStartElement("pic", "cNvPr");
        writer.WriteAttributeString("id", "0");
        writer.WriteAttributeString("name", picName);
        writer.WriteAttributeString("descr", descr);
        writer.WriteEndElement();
        writer.WriteStartElement("pic", "cNvPicPr");
        writer.WriteStartElement("a", "picLocks");
        writer.WriteAttributeString("noChangeAspect", "1");
        writer.WriteEndElement();
        writer.WriteEndElement();
        writer.WriteEndElement(); // nvPicPr

        writer.WriteStartElement("pic", "blipFill");
        writer.WriteStartElement("a", "blip");
        writer.WriteAttributeString("r", "embed", picture.RelationId);
        writer.WriteEndElement();
        writer.WriteStartElement("a", "stretch");
        writer.WriteStartElement("a", "fillRect");
        writer.WriteEndElement();
        writer.WriteEndElement();
        writer.WriteEndElement(); // blipFill

        writer.WriteStartElement("pic", "spPr");
        writer.WriteStartElement("a", "xfrm");
        if (picture.RotationAttribute != 0)
        {
            writer.WriteAttributeString("rot", picture.RotationAttribute.ToString(CultureInfo.InvariantCulture));
        }
        writer.WriteStartElement("a", "off");
        writer.WriteAttributeString("x", "0");
        writer.WriteAttributeString("y", "0");
        writer.WriteEndElement();
        writer.WriteStartElement("a", "ext");
        writer.WriteAttributeString("cx", cx);
        writer.WriteAttributeString("cy", cy);
        writer.WriteEndElement();
        writer.WriteEndElement(); // xfrm
        writer.WriteStartElement("a", "prstGeom");
        writer.WriteAttributeString("prst", "rect");
        writer.WriteStartElement("a", "avLst");
        writer.WriteEndElement();
        writer.WriteEndElement(); // prstGeom
        writer.WriteEndElement(); // spPr

        writer.WriteEndElement(); // pic:pic
        writer.WriteEndElement(); // a:graphicData
        writer.WriteEndElement(); // a:graphic

        writer.WriteEndElement(); // wp:inline
        writer.WriteEndElement(); // w:drawing
        writer.WriteEndElement(); // w:r
    }

    private static void WriteEntry(ZipArchive archive, string name, Action<PoiXmlWriter> write)
    {
        var entry = archive.CreateEntry(name, CompressionLevel.Optimal);
        using var entryStream = entry.Open();
        using var textWriter = new StreamWriter(entryStream, new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
        using var writer = PoiXmlWriterFactory.CreateForOoxmlPackagePart(textWriter, name);
        write(writer);
    }

    private static void WriteBinaryEntry(ZipArchive archive, string name, byte[] data)
    {
        var entry = archive.CreateEntry(name, CompressionLevel.Optimal);
        using var entryStream = entry.Open();
        entryStream.Write(data, 0, data.Length);
    }

    private static void WriteDefault(PoiXmlWriter writer, string extension, string contentType)
    {
        writer.WriteStartElement("Default");
        writer.WriteAttributeString("Extension", extension);
        writer.WriteAttributeString("ContentType", contentType);
        writer.WriteEndElement();
    }

    private static void WriteOverride(PoiXmlWriter writer, string partName, string contentType)
    {
        writer.WriteStartElement("Override");
        writer.WriteAttributeString("PartName", partName);
        writer.WriteAttributeString("ContentType", contentType);
        writer.WriteEndElement();
    }

    private static void WriteRelationship(PoiXmlWriter writer, string id, string target, string type)
    {
        writer.WriteStartElement("Relationship");
        writer.WriteAttributeString("Id", id);
        writer.WriteAttributeString("Target", target);
        writer.WriteAttributeString("Type", type);
        writer.WriteEndElement();
    }

    private static void WriteRelationship(PoiXmlWriter writer, string id, string target, string type, string targetMode)
    {
        writer.WriteStartElement("Relationship");
        writer.WriteAttributeString("Id", id);
        writer.WriteAttributeString("Target", target);
        writer.WriteAttributeString("Type", type);
        writer.WriteAttributeString("TargetMode", targetMode);
        writer.WriteEndElement();
    }

    private void WriteNumbering(PoiXmlWriter writer)
    {
        writer.WriteStartElement("w", "numbering");
        writer.WriteAttributeString("xmlns:w", NsW);
        foreach (var (absId, fmt, text, start) in _abstractNums)
        {
            writer.WriteStartElement("w", "abstractNum");
            writer.WriteAttributeString("w", "abstractNumId", absId.ToString(CultureInfo.InvariantCulture));
            writer.WriteStartElement("w", "lvl");
            writer.WriteAttributeString("w", "ilvl", "0");
            writer.WriteAttributeString("w", "tplc", "FFFFFFFF");
            writer.WriteStartElement("w", "start");
            writer.WriteAttributeString("w", "val", start.ToString(CultureInfo.InvariantCulture));
            writer.WriteEndElement();
            writer.WriteStartElement("w", "numFmt");
            writer.WriteAttributeString("w", "val", fmt);
            writer.WriteEndElement();
            writer.WriteStartElement("w", "lvlText");
            writer.WriteAttributeString("w", "val", text);
            writer.WriteEndElement();
            writer.WriteEndElement(); // lvl
            writer.WriteEndElement(); // abstractNum
        }
        foreach (var (numId, abstractNumId) in _numInstances)
        {
            writer.WriteStartElement("w", "num");
            writer.WriteAttributeString("w", "numId", numId.ToString(CultureInfo.InvariantCulture));
            writer.WriteStartElement("w", "abstractNumId");
            writer.WriteAttributeString("w", "val", abstractNumId.ToString(CultureInfo.InvariantCulture));
            writer.WriteEndElement();
            writer.WriteEndElement(); // num
        }
        writer.WriteEndElement(); // numbering
    }

    private void ReadNumbering(ZipArchive archive)
    {
        var numberingEntry = archive.GetEntry("word/numbering.xml");
        if (numberingEntry is null) return;

        using var stream = numberingEntry.Open();
        using var reader = XmlReader.Create(stream, new XmlReaderSettings { IgnoreWhitespace = false });

        int? currentAbstractNumId = null;
        int? currentNumId = null;

        while (reader.Read())
        {
            if (reader.NodeType == XmlNodeType.Element && reader.NamespaceURI == NsW)
            {
                switch (reader.LocalName)
                {
                    case "abstractNum":
                        var absIdAttr = reader.GetAttribute("w:abstractNumId");
                        if (absIdAttr is not null && int.TryParse(absIdAttr, out var absId))
                        {
                            currentAbstractNumId = absId;
                            if (absId >= _nextAbstractNumId)
                                _nextAbstractNumId = absId + 1;
                        }
                        break;
                    case "num":
                        var numIdAttr = reader.GetAttribute("w:numId");
                        if (numIdAttr is not null && int.TryParse(numIdAttr, out var nid))
                        {
                            currentNumId = nid;
                            if (nid >= _nextNumId)
                                _nextNumId = nid + 1;
                        }
                        break;
                    case "abstractNumId":
                        if (currentNumId is not null)
                        {
                            var absNumRef = reader.GetAttribute("w:val");
                            if (absNumRef is not null && int.TryParse(absNumRef, out var absNumVal))
                            {
                                _numInstances.Add((currentNumId.Value, absNumVal));
                            }
                        }
                        break;
                    case "numFmt":
                        if (currentAbstractNumId is not null)
                        {
                            var fmtVal = reader.GetAttribute("w:val");
                            _abstractNums.Add((currentAbstractNumId.Value,
                                fmtVal ?? "decimal", "\u2022", 1));
                        }
                        break;
                }
            }
            if (reader.NodeType == XmlNodeType.EndElement && reader.NamespaceURI == NsW)
            {
                if (reader.LocalName == "abstractNum") currentAbstractNumId = null;
                if (reader.LocalName == "num") currentNumId = null;
            }
        }
    }
}
