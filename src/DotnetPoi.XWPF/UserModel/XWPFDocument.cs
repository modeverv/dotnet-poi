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

    private readonly List<XWPFParagraph> _paragraphs = new();
    private readonly List<XWPFPictureData> _pictures = new();
    private long _nextDrawingId = 1;

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
        if (_pictures.Count > 0)
        {
            WriteEntry(archive, "word/_rels/document.xml.rels", WriteDocumentRelationships);
        }
        foreach (var pic in _pictures)
        {
            WriteBinaryEntry(archive, $"word/media/{pic.getFileName()}", pic.Data);
        }
    }

    public void close() { }

    public void Dispose() => close();

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

        using var archive = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: true);

        ReadPictures(archive);
        ReadDocument(archive);
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
                                currentParagraph = createParagraph();
                                break;
                            case "r" when !inRPr:
                                if (currentParagraph is not null)
                                    currentRun = currentParagraph.createRun();
                                break;
                            case "rPr":
                                inRPr = true;
                                break;
                            case "b" when inRPr:
                                currentRun?.setBold(true);
                                break;
                            case "i" when inRPr:
                                currentRun?.setItalic(true);
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
                                break;
                            case "r" when !inRPr:
                                currentRun = null;
                                break;
                            case "rPr":
                                inRPr = false;
                                break;
                            case "drawing" when inDrawing:
                                inDrawing = false;
                                inInline = false;
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
        writer.WriteStartDocument("UTF-8", standalone: true);
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
        WriteOverride(writer, "/word/document.xml",
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml");
        writer.WriteEndElement();
    }

    private static void WriteRootRelationships(PoiXmlWriter writer)
    {
        writer.WriteStartDocument("UTF-8", standalone: true);
        writer.WriteStartElement("Relationships");
        writer.WriteAttributeString("xmlns", "http://schemas.openxmlformats.org/package/2006/relationships");
        WriteRelationship(writer, "rId1", "word/document.xml",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument");
        writer.WriteEndElement();
    }

    private void WriteDocumentRelationships(PoiXmlWriter writer)
    {
        writer.WriteStartDocument("UTF-8", standalone: true);
        writer.WriteStartElement("Relationships");
        writer.WriteAttributeString("xmlns", "http://schemas.openxmlformats.org/package/2006/relationships");
        foreach (var pic in _pictures)
        {
            WriteRelationship(writer, $"rId{pic.Index}", $"media/{pic.getFileName()}",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image");
        }
        writer.WriteEndElement();
    }

    private void WriteDocument(PoiXmlWriter writer)
    {
        writer.WriteStartDocument("UTF-8", standalone: true);
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
        writer.WriteStartElement("w", "sectPr");
        writer.WriteEndElement();
        writer.WriteEndElement(); // body
        writer.WriteEndElement(); // document
    }

    private static void WriteParagraph(PoiXmlWriter writer, XWPFParagraph para)
    {
        writer.WriteStartElement("w", "p");
        foreach (var run in para.Runs)
        {
            WriteRun(writer, run);
        }
        writer.WriteEndElement();
    }

    private static void WriteRun(PoiXmlWriter writer, XWPFRun run)
    {
        if (run.TextValue is not null)
        {
            writer.WriteStartElement("w", "r");
            if (run.Bold || run.Italic)
            {
                writer.WriteStartElement("w", "rPr");
                if (run.Bold)
                {
                    writer.WriteStartElement("w", "b");
                    writer.WriteEndElement();
                }
                if (run.Italic)
                {
                    writer.WriteStartElement("w", "i");
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
        using var writer = new PoiXmlWriter(textWriter);
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
}
