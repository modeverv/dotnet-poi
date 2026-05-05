using System.Globalization;
using System.IO.Compression;
using System.Text;
using System.Xml;
using DotnetPoi.SS.Xml;

namespace DotnetPoi.XWPF.UserModel;

public sealed class XWPFDocument : IDisposable
{
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

    /// <summary>True if this document was loaded from docm (has a vbaProject.bin).</summary>
    public bool HasMacros => _vbaProjectBin != null;

    public bool isMacroEnabled() => HasMacros;

    // rId1 is always reserved for settings.xml; image rIds start at rId{Index + 1}.
    private const int ImageRelIdOffset = 1;

    private readonly List<XWPFParagraph> _paragraphs = new();
    private readonly List<XWPFPictureData> _pictures = new();
    private readonly List<XWPFTable> _tables = new();
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

    public XWPFDocument() { }

    public XWPFDocument(Stream stream)
    {
        ArgumentNullException.ThrowIfNull(stream);
        Load(stream);
    }

    public XWPFParagraph createParagraph()
    {
        var paragraph = new XWPFParagraph(this);
        _paragraphs.Add(paragraph);
        return paragraph;
    }

    public IReadOnlyList<XWPFParagraph> getParagraphs() => _paragraphs;

    public XWPFTable createTable()
    {
        var table = new XWPFTable(this);
        _tables.Add(table);
        return table;
    }

    public IReadOnlyList<XWPFTable> getTables() => _tables;

    public IReadOnlyList<XWPFPictureData> getAllPictures() => _pictures;

    /// <summary>Returns 0-based index of added picture data.</summary>
    public int addPicture(byte[] pictureData, int format)
    {
        ArgumentNullException.ThrowIfNull(pictureData);
        var data = AddPictureData(pictureData, format);
        return data.Index - 1;
    }

    public void write(Stream stream)
    {
        ArgumentNullException.ThrowIfNull(stream);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true);
        WriteEntry(archive, "[Content_Types].xml", WriteContentTypes);
        WriteEntry(archive, "_rels/.rels", WriteRootRelationships);
        WriteEntry(archive, "word/document.xml", WriteDocument);
        WriteEntry(archive, "word/_rels/document.xml.rels", WriteDocumentRelationships);
        WriteEntry(archive, "word/settings.xml", WriteSettings);
        if (_numInstances.Count > 0)
            WriteEntry(archive, "word/numbering.xml", WriteNumbering);
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
        ArgumentNullException.ThrowIfNull(vbaProjectData);
        _vbaProjectBin = vbaProjectData.ToArray();
        _vbaDataXml = null;
        _vbaProjectBinRels = null;
    }

    public void setVBAProject(Stream vbaProjectStream)
    {
        ArgumentNullException.ThrowIfNull(vbaProjectStream);
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
        _pictures.Clear();
        _nextDrawingId = 1;
        _vbaProjectBin = null;
        _vbaDataXml = null;
        _vbaProjectBinRels = null;

        using var archive = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: true);

        ReadPictures(archive);
        ReadDocument(archive);
        ReadVbaFiles(archive);
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

    private void ReadDocument(ZipArchive archive)
    {
        // Build relId → picture-index map from document relationships
        var picIndexByRelId = BuildPictureRelMap(archive);

        var docEntry = archive.GetEntry("word/document.xml");
        if (docEntry is null) return;

        using var docStream = docEntry.Open();
        using var reader = XmlReader.Create(docStream, new XmlReaderSettings { IgnoreWhitespace = false });

        XWPFParagraph? currentParagraph = null;
        XWPFRun? currentRun = null;
        XWPFTable? currentTable = null;
        XWPFTableRow? currentRow = null;
        XWPFTableCell? currentTableCell = null;
        bool inPPr = false;
        bool inRPr = false;

        // Inline image parsing state
        bool inDrawing = false;
        long wpInlineExtCx = 0, wpInlineExtCy = 0;
        string? wpDocPrDescr = null;
        long wpDocPrId = 0;
        string? blipEmbed = null;
        int xfrmRot = 0;
        bool inInline = false;

        while (reader.Read())
        {
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
                            case "r" when !inRPr:
                                if (currentParagraph is not null)
                                    currentRun = currentParagraph.createRun();
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
                            case "t" when currentRun is not null && !inRPr:
                                currentRun.setText(reader.ReadElementContentAsString());
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
                                break;
                            case "r" when !inRPr:
                                currentRun = null;
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
                            case "tbl":
                                currentTable = null;
                                break;
                            case "tr":
                                currentRow = null;
                                break;
                            case "tc":
                                currentTableCell = null;
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
                && int.TryParse(mediaFile["image".Length..], out var oneBasedIndex))
            {
                map[relId] = oneBasedIndex - 1;
            }
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
        WriteDefault(writer, "xml", "application/xml");
        // docm uses macroEnabled content type; docx uses the standard one.
        WriteOverride(writer, "/word/document.xml",
            _vbaProjectBin != null ? ContentTypeDocm : ContentTypeDocx);
        WriteOverride(writer, "/word/settings.xml", ContentTypeSettings);
        if (_numInstances.Count > 0)
            WriteOverride(writer, "/word/numbering.xml", ContentTypeNumbering);
        if (_vbaProjectBin != null)
        {
            WriteOverride(writer, "/word/vbaProject.bin", ContentTypeVbaProject);
            if (_vbaDataXml != null)
                WriteOverride(writer, "/word/vbaData.xml", ContentTypeVbaData);
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
        // rId1 is always settings; images use rId{Index + ImageRelIdOffset}
        WriteRelationship(writer, "rId1", "settings.xml", RelTypeSettings);
        foreach (var pic in _pictures)
        {
            WriteRelationship(writer, $"rId{pic.Index + ImageRelIdOffset}", $"media/{pic.getFileName()}",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image");
        }
        // docm: vbaProject relationship after images
        if (_vbaProjectBin != null)
        {
            var vbaRelId = $"rId{_pictures.Count + 2}";
            WriteRelationship(writer, vbaRelId, "vbaProject.bin", RelTypeVbaProject);
        }
        // numbering relationship after images
        if (_numInstances.Count > 0)
        {
            var numRelId = _vbaProjectBin != null
                ? $"rId{_pictures.Count + 3}"
                : $"rId{_pictures.Count + 2}";
            WriteRelationship(writer, numRelId, "numbering.xml", RelTypeNumbering);
        }
        writer.WriteEndElement();
    }

    private static void WriteSettings(PoiXmlWriter writer)
    {
        writer.WriteStartElement("w", "settings");
        writer.WriteAttributeString("xmlns:w", NsW);
        writer.WriteEndElement();
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
        foreach (var para in _paragraphs)
        {
            WriteParagraph(writer, para);
        }
        foreach (var table in _tables)
        {
            WriteTable(writer, table);
        }
        writer.WriteStartElement("w", "sectPr");
        writer.WriteEndElement();
        writer.WriteEndElement(); // body
        writer.WriteEndElement(); // document
    }

    private static void WriteParagraph(PoiXmlWriter writer, XWPFParagraph para)
    {
        writer.WriteStartElement("w", "p");

        if (para.Alignment is not null || para.IndentLeft != 0
            || para.IndentRight != 0 || para.IndentFirstLine != 0 || para.IndentHanging != 0
            || para.SpacingBefore != 0 || para.SpacingAfter != 0 || para.SpacingBetween != 0
            || para.NumId is not null)
        {
            writer.WriteStartElement("w", "pPr");
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
            writer.WriteEndElement(); // pPr
        }

        foreach (var run in para.Runs)
        {
            WriteRun(writer, run);
        }
        writer.WriteEndElement();
    }

    private static void WriteTable(PoiXmlWriter writer, XWPFTable table)
    {
        writer.WriteStartElement("w", "tbl");
        writer.WriteStartElement("w", "tblPr");
        writer.WriteStartElement("w", "tblW");
        writer.WriteAttributeString("w", "w", "5000");
        writer.WriteAttributeString("w", "type", "dxa");
        writer.WriteEndElement();
        writer.WriteEndElement(); // tblPr
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
            foreach (var cell in row.Cells)
            {
                writer.WriteStartElement("w", "tc");
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

    private static void WriteRun(PoiXmlWriter writer, XWPFRun run)
    {
        if (run.TextValue is not null)
        {
            writer.WriteStartElement("w", "r");
            if (run.Bold || run.Italic || run.Underline || run.Strike
                || run.FontName is not null || run.FontSize > 0 || run.Color is not null)
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
                writer.WriteEndElement(); // rPr
            }
            writer.WriteStartElement("w", "t");
            writer.WriteAttributeString("xml:space", "preserve");
            writer.WriteString(run.TextValue);
            writer.WriteEndElement(); // t
            writer.WriteEndElement(); // r
        }
        foreach (var picture in run.Pictures)
        {
            WriteInlinePicture(writer, picture);
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
