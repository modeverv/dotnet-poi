using System.Globalization;
using System.IO.Compression;
using System.Text;
using System.Xml;
using DotnetPoi.POIFS.Crypt;
using DotnetPoi.SS.Xml;

namespace DotnetPoi.XSLF.UserModel;

/// <summary>
/// High-level representation of a PPTX presentation.
/// Ported from org.apache.poi.xslf.usermodel.XMLSlideShow.
/// </summary>
public sealed class XMLSlideShow : IDisposable
{
    // Default slide size: 10 × 7.5 inches in EMU
    public const long DefaultSlideCx = 9_144_000L;
    public const long DefaultSlideCy = 6_858_000L;

    // Namespace URIs (used by XmlReader on the read path)
    private const string NsP   = "http://schemas.openxmlformats.org/presentationml/2006/main";
    private const string NsA   = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private const string NsR   = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    private const string RelTypeOfficeDoc  = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
    private const string RelTypeSlide      = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide";
    private const string RelTypeSlideMaster= "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster";
    private const string RelTypeSlideLayout= "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout";
    private const string RelTypeImage      = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
    private const string RelTypePresProps  = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/presProps";
    private const string RelTypeTheme      = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme";
    private const string RelTypeTableStyles= "http://schemas.openxmlformats.org/officeDocument/2006/relationships/tableStyles";

    private const string ContentTypePresentation = "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml";
    private const string ContentTypePptm        = "application/vnd.ms-powerpoint.presentation.macroEnabled.main+xml";
    private const string ContentTypeVbaProject  = "application/vnd.ms-office.vbaProject";
    private const string RelTypeVbaProject      = "http://schemas.microsoft.com/office/2006/relationships/vbaProject";

    /// <summary>True if this presentation was loaded from pptm (has a vbaProject.bin).</summary>
    public bool HasMacros => _vbaProjectBin != null;
    public bool isMacroEnabled() => HasMacros;
    private const string ContentTypeSlide        = "application/vnd.openxmlformats-officedocument.presentationml.slide+xml";
    private const string ContentTypeSlideMaster  = "application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml";
    private const string ContentTypeSlideLayout  = "application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml";
    private const string ContentTypePresProps    = "application/vnd.openxmlformats-officedocument.presentationml.presProps+xml";
    private const string ContentTypeTheme        = "application/vnd.openxmlformats-officedocument.theme+xml";
    private const string ContentTypeTableStyles  = "application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml";

    // Presentation.xml.rels has fixed rels rId1-rId5; slides start at rId6.
    private const int SlideRelIdOffset = 6;

    // Standard Office Theme (same as Apache POI default)
    private static readonly byte[] OfficeThemeBytes = Encoding.UTF8.GetBytes(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
        "<a:theme xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" name=\"Office Theme\">" +
        "<a:themeElements>" +
        "<a:clrScheme name=\"Office\">" +
        "<a:dk1><a:sysClr val=\"windowText\" lastClr=\"000000\"/></a:dk1>" +
        "<a:lt1><a:sysClr val=\"window\" lastClr=\"FFFFFF\"/></a:lt1>" +
        "<a:dk2><a:srgbClr val=\"1F497D\"/></a:dk2>" +
        "<a:lt2><a:srgbClr val=\"EEECE1\"/></a:lt2>" +
        "<a:accent1><a:srgbClr val=\"4F81BD\"/></a:accent1>" +
        "<a:accent2><a:srgbClr val=\"C0504D\"/></a:accent2>" +
        "<a:accent3><a:srgbClr val=\"9BBB59\"/></a:accent3>" +
        "<a:accent4><a:srgbClr val=\"8064A2\"/></a:accent4>" +
        "<a:accent5><a:srgbClr val=\"4BACC6\"/></a:accent5>" +
        "<a:accent6><a:srgbClr val=\"F79646\"/></a:accent6>" +
        "<a:hlink><a:srgbClr val=\"0000FF\"/></a:hlink>" +
        "<a:folHlink><a:srgbClr val=\"800080\"/></a:folHlink>" +
        "</a:clrScheme>" +
        "<a:fontScheme name=\"Office\">" +
        "<a:majorFont><a:latin typeface=\"Calibri\"/><a:ea typeface=\"\"/><a:cs typeface=\"\"/></a:majorFont>" +
        "<a:minorFont><a:latin typeface=\"Calibri\"/><a:ea typeface=\"\"/><a:cs typeface=\"\"/></a:minorFont>" +
        "</a:fontScheme>" +
        "<a:fmtScheme name=\"Office\">" +
        "<a:fillStyleLst>" +
        "<a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill>" +
        "<a:gradFill rotWithShape=\"1\"><a:gsLst>" +
        "<a:gs pos=\"0\"><a:schemeClr val=\"phClr\"><a:tint val=\"50000\"/><a:satMod val=\"300000\"/></a:schemeClr></a:gs>" +
        "<a:gs pos=\"35000\"><a:schemeClr val=\"phClr\"><a:tint val=\"37000\"/><a:satMod val=\"300000\"/></a:schemeClr></a:gs>" +
        "<a:gs pos=\"100000\"><a:schemeClr val=\"phClr\"><a:tint val=\"15000\"/><a:satMod val=\"350000\"/></a:schemeClr></a:gs>" +
        "</a:gsLst><a:lin ang=\"16200000\" scaled=\"1\"/></a:gradFill>" +
        "<a:gradFill rotWithShape=\"1\"><a:gsLst>" +
        "<a:gs pos=\"0\"><a:schemeClr val=\"phClr\"><a:shade val=\"51000\"/><a:satMod val=\"130000\"/></a:schemeClr></a:gs>" +
        "<a:gs pos=\"80000\"><a:schemeClr val=\"phClr\"><a:shade val=\"93000\"/><a:satMod val=\"130000\"/></a:schemeClr></a:gs>" +
        "<a:gs pos=\"100000\"><a:schemeClr val=\"phClr\"><a:shade val=\"94000\"/><a:satMod val=\"135000\"/></a:schemeClr></a:gs>" +
        "</a:gsLst><a:lin ang=\"16200000\" scaled=\"0\"/></a:gradFill>" +
        "</a:fillStyleLst>" +
        "<a:lnStyleLst>" +
        "<a:ln w=\"9525\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\"><a:solidFill><a:schemeClr val=\"phClr\"><a:shade val=\"95000\"/><a:satMod val=\"105000\"/></a:schemeClr></a:solidFill><a:prstDash val=\"solid\"/></a:ln>" +
        "<a:ln w=\"25400\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\"><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:prstDash val=\"solid\"/></a:ln>" +
        "<a:ln w=\"38100\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\"><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:prstDash val=\"solid\"/></a:ln>" +
        "</a:lnStyleLst>" +
        "<a:effectStyleLst>" +
        "<a:effectStyle><a:effectLst><a:outerShdw blurRad=\"40000\" dist=\"20000\" dir=\"5400000\" rotWithShape=\"0\"><a:srgbClr val=\"000000\"><a:alpha val=\"38000\"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle>" +
        "<a:effectStyle><a:effectLst><a:outerShdw blurRad=\"40000\" dist=\"23000\" dir=\"5400000\" rotWithShape=\"0\"><a:srgbClr val=\"000000\"><a:alpha val=\"35000\"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle>" +
        "<a:effectStyle><a:effectLst><a:outerShdw blurRad=\"40000\" dist=\"23000\" dir=\"5400000\" rotWithShape=\"0\"><a:srgbClr val=\"000000\"><a:alpha val=\"35000\"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle>" +
        "</a:effectStyleLst>" +
        "<a:bgFillStyleLst>" +
        "<a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill>" +
        "<a:gradFill rotWithShape=\"1\"><a:gsLst>" +
        "<a:gs pos=\"0\"><a:schemeClr val=\"phClr\"><a:tint val=\"40000\"/><a:satMod val=\"350000\"/></a:schemeClr></a:gs>" +
        "<a:gs pos=\"40000\"><a:schemeClr val=\"phClr\"><a:tint val=\"45000\"/><a:shade val=\"99000\"/><a:satMod val=\"350000\"/></a:schemeClr></a:gs>" +
        "<a:gs pos=\"100000\"><a:schemeClr val=\"phClr\"><a:shade val=\"20000\"/><a:satMod val=\"255000\"/></a:schemeClr></a:gs>" +
        "</a:gsLst><a:path path=\"circle\"><a:fillToRect l=\"50000\" t=\"-80000\" r=\"50000\" b=\"180000\"/></a:path></a:gradFill>" +
        "<a:gradFill rotWithShape=\"1\"><a:gsLst>" +
        "<a:gs pos=\"0\"><a:schemeClr val=\"phClr\"><a:tint val=\"80000\"/><a:satMod val=\"300000\"/></a:schemeClr></a:gs>" +
        "<a:gs pos=\"100000\"><a:schemeClr val=\"phClr\"><a:shade val=\"30000\"/><a:satMod val=\"200000\"/></a:schemeClr></a:gs>" +
        "</a:gsLst><a:path path=\"circle\"><a:fillToRect l=\"50000\" t=\"50000\" r=\"50000\" b=\"50000\"/></a:path></a:gradFill>" +
        "</a:bgFillStyleLst>" +
        "</a:fmtScheme>" +
        "</a:themeElements>" +
        "<a:objectDefaults/>" +
        "<a:extraClrSchemeLst/>" +
        "</a:theme>");

    private readonly List<XSLFSlide>       _slides   = new();
    private readonly List<XSLFPictureData> _pictures = new();
    private long _slideCx = DefaultSlideCx;
    private long _slideCy = DefaultSlideCy;

    // pptm support: opaque VBA binary preserved byte-for-byte.
    private byte[]? _vbaProjectBin;

    // Unknown-part preservation: non-model ZIP entries stored verbatim during Load and re-emitted during write.
    private Dictionary<string, byte[]> _preservedEntries = new(StringComparer.OrdinalIgnoreCase);

    public XMLSlideShow() { }

    public XMLSlideShow(Stream stream)
    {
        ArgumentNullException.ThrowIfNull(stream);
        Load(stream);
    }

    // ----- public API -----

    public XSLFSlide createSlide()
    {
        var slide = new XSLFSlide();
        _slides.Add(slide);
        return slide;
    }

    /// <summary>Adds picture bytes and returns the 0-based picture index.</summary>
    public int addPicture(byte[] data, int format)
    {
        ArgumentNullException.ThrowIfNull(data);
        var existing = _pictures.FirstOrDefault(p => p.Format == format && p.Data.SequenceEqual(data));
        if (existing is not null) return existing.Index - 1;
        var pic = new XSLFPictureData(data, format, _pictures.Count + 1);
        _pictures.Add(pic);
        return pic.Index - 1;
    }

    /// <summary>
    /// Creates a picture shape on the given slide referencing the picture at the 0-based index.
    /// </summary>
    public XSLFPictureShape createPicture(XSLFSlide slide, int pictureIndex)
    {
        ArgumentNullException.ThrowIfNull(slide);
        if ((uint)pictureIndex >= (uint)_pictures.Count)
            throw new ArgumentOutOfRangeException(nameof(pictureIndex));
        var data  = _pictures[pictureIndex];
        // rId1 is reserved for the slideLayout relationship in each slide's rels file;
        // image relationships start at rId2 = rId{Index + 1}.
        var relId = $"rId{data.Index + 1}";
        return slide.CreatePicture(data, relId);
    }

    public IReadOnlyList<XSLFSlide>       getSlides()      => _slides;
    public IReadOnlyList<XSLFPictureData> getPictureData() => _pictures;

    public long getSlideCx() => _slideCx;
    public long getSlideCy() => _slideCy;

    public void setSlideSize(long cx, long cy)
    {
        _slideCx = cx;
        _slideCy = cy;
    }

    public void write(Stream stream)
    {
        ArgumentNullException.ThrowIfNull(stream);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true);

        // Emit preserved (non-model) entries first, so model entries overwrite them
        foreach (var kv in _preservedEntries)
            WriteBinaryEntry(archive, kv.Key, kv.Value);

        WriteEntry(archive, "[Content_Types].xml",                           WriteContentTypes);
        WriteEntry(archive, "_rels/.rels",                                   WriteRootRelationships);
        WriteEntry(archive, "ppt/presentation.xml",                          WritePresentation);
        WriteEntry(archive, "ppt/_rels/presentation.xml.rels",               WritePresentationRels);
        WriteEntry(archive, "ppt/presProps.xml",                             WritePresProps);
        WriteEntry(archive, "ppt/tableStyles.xml",                           WriteTableStyles);
        WriteBinaryEntry(archive, "ppt/theme/theme1.xml",                    OfficeThemeBytes);
        WriteEntry(archive, "ppt/slideMasters/slideMaster1.xml",             WriteSlideMasterStub);
        WriteEntry(archive, "ppt/slideMasters/_rels/slideMaster1.xml.rels",  WriteSlideMasterRels);
        WriteEntry(archive, "ppt/slideLayouts/slideLayout1.xml",             WriteSlideLayoutStub);
        WriteEntry(archive, "ppt/slideLayouts/_rels/slideLayout1.xml.rels",  WriteSlideLayoutRels);

        for (int i = 0; i < _slides.Count; i++)
        {
            int slideNum = i + 1;
            var slide    = _slides[i];
            WriteEntry(archive, $"ppt/slides/slide{slideNum}.xml",            w => WriteSlide(w, slide));
            WriteEntry(archive, $"ppt/slides/_rels/slide{slideNum}.xml.rels", w => WriteSlideRels(w, slide));
        }

        foreach (var pic in _pictures)
            WriteBinaryEntry(archive, $"ppt/media/{pic.getFileName()}", pic.Data);

        if (_vbaProjectBin != null)
            WriteBinaryEntry(archive, "ppt/vbaProject.bin", _vbaProjectBin);
    }

    public void writeEncrypted(Stream stream, string password)
    {
        ArgumentNullException.ThrowIfNull(stream);
        ArgumentNullException.ThrowIfNull(password);
        using var package = new MemoryStream();
        write(package);

        var info = new EncryptionInfo(EncryptionMode.agile);
        info.Encryptor.confirmPassword(password);
        info.Encryptor.encryptPackage(package.ToArray(), stream);
    }

    public void close() { }
    public void Dispose() => close();

    public void setVBAProject(byte[] vbaProjectData)
    {
        ArgumentNullException.ThrowIfNull(vbaProjectData);
        _vbaProjectBin = vbaProjectData.ToArray();
    }

    public void setVBAProject(Stream vbaProjectStream)
    {
        ArgumentNullException.ThrowIfNull(vbaProjectStream);
        using var ms = new MemoryStream();
        vbaProjectStream.CopyTo(ms);
        _vbaProjectBin = ms.ToArray();
    }

    // ----- write helpers -----

    private void WriteContentTypes(PoiXmlWriter w)
    {
        w.WriteStartElement("Types");
        w.WriteAttributeString("xmlns", "http://schemas.openxmlformats.org/package/2006/content-types");

        WriteDefault(w, "rels", "application/vnd.openxmlformats-package.relationships+xml");
        WriteDefault(w, "xml",  "application/xml");
        foreach (var ext in _pictures
            .Select(p => p.Extension).Distinct(StringComparer.Ordinal)
            .OrderBy(e => e, StringComparer.Ordinal))
        {
            WriteDefault(w, ext, _pictures.First(p => p.Extension == ext).ContentType);
        }

        // pptm uses macroEnabled content type; pptx uses the standard one.
        WriteOverride(w, "/ppt/presentation.xml", _vbaProjectBin != null ? ContentTypePptm : ContentTypePresentation);
        if (_vbaProjectBin != null)
            WriteOverride(w, "/ppt/vbaProject.bin", ContentTypeVbaProject);
        WriteOverride(w, "/ppt/presProps.xml",                   ContentTypePresProps);
        WriteOverride(w, "/ppt/tableStyles.xml",                 ContentTypeTableStyles);
        WriteOverride(w, "/ppt/theme/theme1.xml",                ContentTypeTheme);
        WriteOverride(w, "/ppt/slideMasters/slideMaster1.xml",   ContentTypeSlideMaster);
        WriteOverride(w, "/ppt/slideLayouts/slideLayout1.xml",   ContentTypeSlideLayout);
        for (int i = 0; i < _slides.Count; i++)
            WriteOverride(w, $"/ppt/slides/slide{i + 1}.xml", ContentTypeSlide);

        w.WriteEndElement();
    }

    private static void WriteRootRelationships(PoiXmlWriter w)
    {
        w.WriteStartElement("Relationships");
        w.WriteAttributeString("xmlns", "http://schemas.openxmlformats.org/package/2006/relationships");
        WriteRelationship(w, "rId1", "ppt/presentation.xml", RelTypeOfficeDoc);
        w.WriteEndElement();
    }

    private void WritePresentation(PoiXmlWriter w)
    {
        w.WriteStartElement("p", "presentation");
        w.WriteAttributeString("xmlns:p", NsP);
        w.WriteAttributeString("xmlns:a", NsA);
        w.WriteAttributeString("xmlns:r", NsR);

        w.WriteStartElement("p", "sldMasterIdLst");
        w.WriteStartElement("p", "sldMasterId");
        w.WriteAttributeString("id", "2147483648");
        w.WriteAttributeString("r", "id", "rId1");
        w.WriteEndElement();
        w.WriteEndElement();

        w.WriteStartElement("p", "sldIdLst");
        for (int i = 0; i < _slides.Count; i++)
        {
            w.WriteStartElement("p", "sldId");
            w.WriteAttributeString("id", (256 + i).ToString(CultureInfo.InvariantCulture));
            // Slides start at rId{SlideRelIdOffset} after the 5 static rels
            w.WriteAttributeString("r", "id", $"rId{SlideRelIdOffset + i}");
            w.WriteEndElement();
        }
        w.WriteEndElement();

        w.WriteStartElement("p", "sldSz");
        w.WriteAttributeString("cx", _slideCx.ToString(CultureInfo.InvariantCulture));
        w.WriteAttributeString("cy", _slideCy.ToString(CultureInfo.InvariantCulture));
        w.WriteEndElement();

        w.WriteStartElement("p", "notesSz");
        w.WriteAttributeString("cx", "6858000");
        w.WriteAttributeString("cy", "9144000");
        w.WriteEndElement();

        w.WriteEndElement(); // p:presentation
    }

    private void WritePresentationRels(PoiXmlWriter w)
    {
        w.WriteStartElement("Relationships");
        w.WriteAttributeString("xmlns", "http://schemas.openxmlformats.org/package/2006/relationships");
        // Fixed rels: rId1–rId5 (match SlideRelIdOffset = 6, so slides start at rId6)
        WriteRelationship(w, "rId1", "slideMasters/slideMaster1.xml", RelTypeSlideMaster);
        WriteRelationship(w, "rId2", "presProps.xml",                  RelTypePresProps);
        WriteRelationship(w, "rId3", "theme/theme1.xml",               RelTypeTheme);
        WriteRelationship(w, "rId4", "tableStyles.xml",                RelTypeTableStyles);
        // rId5 is reserved (matches SlideRelIdOffset - 1 = 5) — unused for now; slides start at rId6
        for (int i = 0; i < _slides.Count; i++)
            WriteRelationship(w, $"rId{SlideRelIdOffset + i}", $"slides/slide{i + 1}.xml", RelTypeSlide);
        // pptm: vbaProject relationship appended after slides
        if (_vbaProjectBin != null)
            WriteRelationship(w, $"rId{SlideRelIdOffset + _slides.Count}", "vbaProject.bin", RelTypeVbaProject);
        w.WriteEndElement();
    }

    private static void WritePresProps(PoiXmlWriter w)
    {
        w.WriteStartElement("p", "presentationPr");
        w.WriteAttributeString("xmlns:a", NsA);
        w.WriteAttributeString("xmlns:r", NsR);
        w.WriteAttributeString("xmlns:p", NsP);
        w.WriteEndElement();
    }

    private static void WriteTableStyles(PoiXmlWriter w)
    {
        w.WriteStartElement("a", "tblStyleLst");
        w.WriteAttributeString("xmlns:a", NsA);
        w.WriteAttributeString("def", "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}");
        w.WriteEndElement();
    }

    private static void WriteSlideMasterStub(PoiXmlWriter w)
    {
        w.WriteStartElement("p", "sldMaster");
        w.WriteAttributeString("xmlns:p", NsP);
        w.WriteAttributeString("xmlns:a", NsA);
        w.WriteAttributeString("xmlns:r", NsR);
        WriteSpTreeStub(w);
        w.WriteStartElement("p", "clrMap");
        foreach (var (k, v) in new[] {
            ("bg1","lt1"),("tx1","dk1"),("bg2","lt2"),("tx2","dk2"),
            ("accent1","accent1"),("accent2","accent2"),("accent3","accent3"),
            ("accent4","accent4"),("accent5","accent5"),("accent6","accent6"),
            ("hlink","hlink"),("folHlink","folHlink") })
        { w.WriteAttributeString(k, v); }
        w.WriteEndElement();
        w.WriteStartElement("p", "sldLayoutIdLst");
        w.WriteStartElement("p", "sldLayoutId");
        w.WriteAttributeString("id", "2147483649");
        w.WriteAttributeString("r", "id", "rId1");
        w.WriteEndElement();
        w.WriteEndElement();
        w.WriteStartElement("p", "txStyles");
        WriteTxStyleEntry(w, "titleStyle");
        WriteTxStyleEntry(w, "bodyStyle");
        WriteTxStyleEntry(w, "otherStyle");
        w.WriteEndElement();
        w.WriteEndElement();
    }

    private static void WriteSlideMasterRels(PoiXmlWriter w)
    {
        w.WriteStartElement("Relationships");
        w.WriteAttributeString("xmlns", "http://schemas.openxmlformats.org/package/2006/relationships");
        WriteRelationship(w, "rId1", "../slideLayouts/slideLayout1.xml", RelTypeSlideLayout);
        WriteRelationship(w, "rId2", "../theme/theme1.xml",              RelTypeTheme);
        w.WriteEndElement();
    }

    private static void WriteSlideLayoutStub(PoiXmlWriter w)
    {
        w.WriteStartElement("p", "sldLayout");
        w.WriteAttributeString("xmlns:p", NsP);
        w.WriteAttributeString("xmlns:a", NsA);
        w.WriteAttributeString("xmlns:r", NsR);
        w.WriteAttributeString("type", "blank");
        w.WriteAttributeString("preserve", "1");
        WriteSpTreeStub(w);
        w.WriteStartElement("p", "clrMapOvr");
        w.WriteStartElement("a", "masterClrMapping");
        w.WriteEndElement();
        w.WriteEndElement();
        w.WriteEndElement();
    }

    private static void WriteSlideLayoutRels(PoiXmlWriter w)
    {
        w.WriteStartElement("Relationships");
        w.WriteAttributeString("xmlns", "http://schemas.openxmlformats.org/package/2006/relationships");
        WriteRelationship(w, "rId1", "../slideMasters/slideMaster1.xml", RelTypeSlideMaster);
        w.WriteEndElement();
    }

    private static void WriteSlide(PoiXmlWriter w, XSLFSlide slide)
    {
        w.WriteStartElement("p", "sld");
        w.WriteAttributeString("xmlns:p", NsP);
        w.WriteAttributeString("xmlns:a", NsA);
        w.WriteAttributeString("xmlns:r", NsR);

        w.WriteStartElement("p", "cSld");
        w.WriteStartElement("p", "spTree");

        w.WriteStartElement("p", "nvGrpSpPr");
        w.WriteStartElement("p", "cNvPr");
        w.WriteAttributeString("id", "1"); w.WriteAttributeString("name", "");
        w.WriteEndElement();
        w.WriteStartElement("p", "cNvGrpSpPr"); w.WriteEndElement();
        w.WriteStartElement("p", "nvPr"); w.WriteEndElement();
        w.WriteEndElement();

        w.WriteStartElement("p", "grpSpPr");
        w.WriteStartElement("a", "xfrm");
        WriteAOff(w, 0, 0); WriteAExt(w, 0, 0);
        WriteAChOff(w, 0, 0); WriteAChExt(w, 0, 0);
        w.WriteEndElement(); // a:xfrm
        w.WriteEndElement(); // p:grpSpPr

        foreach (var shape in slide.getShapes())
            WritePicShape(w, shape);

        foreach (var autoShape in slide.getAutoShapes())
            WriteAutoShape(w, autoShape);

        foreach (var table in slide.getTables())
            WriteTableGraphicFrame(w, table);

        w.WriteEndElement(); // p:spTree
        w.WriteEndElement(); // p:cSld

        w.WriteStartElement("p", "clrMapOvr");
        w.WriteStartElement("a", "masterClrMapping"); w.WriteEndElement();
        w.WriteEndElement();

        w.WriteEndElement(); // p:sld
    }

    private static void WritePicShape(PoiXmlWriter w, XSLFPictureShape shape)
    {
        var id = shape.ShapeId.ToString(CultureInfo.InvariantCulture);
        var cx = shape.AnchorCx.ToString(CultureInfo.InvariantCulture);
        var cy = shape.AnchorCy.ToString(CultureInfo.InvariantCulture);
        var x  = shape.AnchorX.ToString(CultureInfo.InvariantCulture);
        var y  = shape.AnchorY.ToString(CultureInfo.InvariantCulture);

        w.WriteStartElement("p", "pic");

        w.WriteStartElement("p", "nvPicPr");
        w.WriteStartElement("p", "cNvPr");
        w.WriteAttributeString("id", id); w.WriteAttributeString("name", $"Picture {id}");
        w.WriteEndElement();
        w.WriteStartElement("p", "cNvPicPr");
        w.WriteStartElement("a", "picLocks"); w.WriteAttributeString("noChangeAspect", "1"); w.WriteEndElement();
        w.WriteEndElement();
        w.WriteStartElement("p", "nvPr"); w.WriteEndElement();
        w.WriteEndElement(); // nvPicPr

        w.WriteStartElement("p", "blipFill");
        w.WriteStartElement("a", "blip"); w.WriteAttributeString("r", "embed", shape.RelationId); w.WriteEndElement();
        w.WriteStartElement("a", "stretch");
        w.WriteStartElement("a", "fillRect"); w.WriteEndElement();
        w.WriteEndElement();
        w.WriteEndElement(); // blipFill

        w.WriteStartElement("p", "spPr");
        w.WriteStartElement("a", "xfrm");
        if (shape.RotationAttribute != 0) w.WriteAttributeString("rot", shape.RotationAttribute.ToString(CultureInfo.InvariantCulture));
        if (shape.FlipH)                  w.WriteAttributeString("flipH", "1");
        if (shape.FlipV)                  w.WriteAttributeString("flipV", "1");
        WriteAOff(w, shape.AnchorX, shape.AnchorY);
        WriteAExt(w, shape.AnchorCx, shape.AnchorCy);
        w.WriteEndElement(); // a:xfrm
        w.WriteStartElement("a", "prstGeom"); w.WriteAttributeString("prst", "rect");
        w.WriteStartElement("a", "avLst"); w.WriteEndElement();
        w.WriteEndElement();
        w.WriteEndElement(); // p:spPr

        w.WriteEndElement(); // p:pic
    }

    private static void WriteAutoShape(PoiXmlWriter w, XSLFAutoShape shape)
    {
        var id = shape.ShapeId.ToString(CultureInfo.InvariantCulture);
        var cx = shape.AnchorCx.ToString(CultureInfo.InvariantCulture);
        var cy = shape.AnchorCy.ToString(CultureInfo.InvariantCulture);
        var x  = shape.AnchorX.ToString(CultureInfo.InvariantCulture);
        var y  = shape.AnchorY.ToString(CultureInfo.InvariantCulture);

        w.WriteStartElement("p", "sp");

        // nvSpPr
        w.WriteStartElement("p", "nvSpPr");
        w.WriteStartElement("p", "cNvPr");
        w.WriteAttributeString("id", id); w.WriteAttributeString("name", $"TextBox {id}");
        w.WriteEndElement();
        w.WriteStartElement("p", "cNvSpPr"); w.WriteAttributeString("txBox", "1"); w.WriteEndElement();
        w.WriteStartElement("p", "nvPr"); w.WriteEndElement();
        w.WriteEndElement(); // nvSpPr

        // spPr
        w.WriteStartElement("p", "spPr");
        w.WriteStartElement("a", "xfrm");
        if (shape.RotationAttribute != 0) w.WriteAttributeString("rot", shape.RotationAttribute.ToString(CultureInfo.InvariantCulture));
        if (shape.FlipH)                  w.WriteAttributeString("flipH", "1");
        if (shape.FlipV)                  w.WriteAttributeString("flipV", "1");
        WriteAOff(w, shape.AnchorX, shape.AnchorY);
        WriteAExt(w, shape.AnchorCx, shape.AnchorCy);
        w.WriteEndElement(); // a:xfrm
        w.WriteStartElement("a", "prstGeom"); w.WriteAttributeString("prst", "rect");
        w.WriteStartElement("a", "avLst"); w.WriteEndElement();
        w.WriteEndElement();
        w.WriteEndElement(); // p:spPr

        // txBody
        w.WriteStartElement("p", "txBody");
        w.WriteStartElement("a", "bodyPr"); w.WriteEndElement();
        w.WriteStartElement("a", "lstStyle"); w.WriteEndElement();
        foreach (var para in shape.Paragraphs)
        {
            w.WriteStartElement("a", "p");
            foreach (var run in para.Runs)
            {
                w.WriteStartElement("a", "r");
                bool hasRPr = run.Bold || run.Italic || run.Underline || run.Strikethrough
                    || run.FontSize > 0 || run.FontName is not null || run.Color is not null;
                if (hasRPr)
                {
                    w.WriteStartElement("a", "rPr");
                    if (run.Bold)        { w.WriteStartElement("a", "b"); w.WriteEndElement(); }
                    if (run.Italic)      { w.WriteStartElement("a", "i"); w.WriteEndElement(); }
                    if (run.Underline)   { w.WriteStartElement("a", "u"); w.WriteEndElement(); }
                    if (run.Strikethrough) { w.WriteStartElement("a", "strike"); w.WriteEndElement(); }
                    if (run.FontSize > 0)
                    {
                        w.WriteStartElement("a", "sz");
                        w.WriteAttributeString("val", ((int)(run.FontSize * 100)).ToString(CultureInfo.InvariantCulture));
                        w.WriteEndElement();
                    }
                    if (run.FontName is not null)
                    {
                        w.WriteStartElement("a", "latin");
                        w.WriteAttributeString("typeface", run.FontName);
                        w.WriteEndElement();
                    }
                    if (run.Color is not null)
                    {
                        w.WriteStartElement("a", "solidFill");
                        w.WriteStartElement("a", "srgbClr");
                        w.WriteAttributeString("val", run.Color);
                        w.WriteEndElement();
                        w.WriteEndElement();
                    }
                    w.WriteEndElement(); // rPr
                }
                w.WriteStartElement("a", "t");
                w.WriteString(run.Text);
                w.WriteEndElement(); // t
                w.WriteEndElement(); // r
            }
            w.WriteEndElement(); // p
        }
        w.WriteEndElement(); // txBody

        w.WriteEndElement(); // p:sp
    }

    private static void WriteTableGraphicFrame(PoiXmlWriter w, XSLFTable table)
    {
        var id = table.ShapeId.ToString(CultureInfo.InvariantCulture);

        w.WriteStartElement("p", "graphicFrame");

        w.WriteStartElement("p", "nvGraphicFramePr");
        w.WriteStartElement("p", "cNvPr");
        w.WriteAttributeString("id", id);
        w.WriteAttributeString("name", $"Table {id}");
        w.WriteEndElement();
        w.WriteStartElement("p", "cNvGraphicFramePr");
        w.WriteAttributeString("bwMode", "auto");
        w.WriteEndElement();
        w.WriteStartElement("p", "nvPr");
        w.WriteEndElement();
        w.WriteEndElement(); // nvGraphicFramePr

        w.WriteStartElement("p", "xfrm");
        WriteAOff(w, table.AnchorX, table.AnchorY);
        WriteAExt(w, table.AnchorCx, table.AnchorCy);
        w.WriteEndElement(); // xfrm

        w.WriteStartElement("a", "graphic");
        w.WriteAttributeString("xmlns:a", NsA);
        w.WriteStartElement("a", "graphicData");
        w.WriteAttributeString("uri", "http://schemas.openxmlformats.org/drawingml/2006/table");

        w.WriteStartElement("a", "tbl");
        w.WriteStartElement("a", "tblPr");
        w.WriteEndElement(); // tblPr

        w.WriteStartElement("a", "tblGrid");
        foreach (var gridCol in table.GridColWidths)
        {
            w.WriteStartElement("a", "gridCol");
            w.WriteAttributeString("w", gridCol!.ToString(CultureInfo.InvariantCulture));
            w.WriteEndElement();
        }
        w.WriteEndElement(); // tblGrid

        foreach (var row in table.Rows)
        {
            w.WriteStartElement("a", "tr");
            foreach (var cell in row.Cells)
            {
                w.WriteStartElement("a", "tc");
                w.WriteStartElement("a", "txBody");
                w.WriteStartElement("a", "bodyPr");
                w.WriteAttributeString("wrap", "square");
                w.WriteEndElement();
                w.WriteStartElement("a", "lstStyle");
                w.WriteEndElement();
                foreach (var para in cell.Paragraphs)
                {
                    w.WriteStartElement("a", "p");
                    foreach (var run in para.Runs)
                    {
                        w.WriteStartElement("a", "r");
                        bool hasFormatting = run.Bold || run.Italic || run.Underline || run.Strikethrough
                            || run.FontSize > 0 || run.FontName is not null || run.Color is not null;
                        if (hasFormatting)
                        {
                            w.WriteStartElement("a", "rPr");
                            if (run.Bold)        { w.WriteStartElement("a", "b"); w.WriteEndElement(); }
                            if (run.Italic)      { w.WriteStartElement("a", "i"); w.WriteEndElement(); }
                            if (run.Underline)   { w.WriteStartElement("a", "u"); w.WriteEndElement(); }
                            if (run.Strikethrough) { w.WriteStartElement("a", "strike"); w.WriteEndElement(); }
                            if (run.FontSize > 0)
                            {
                                w.WriteStartElement("a", "sz");
                                w.WriteAttributeString("val", ((int)(run.FontSize * 100)).ToString(CultureInfo.InvariantCulture));
                                w.WriteEndElement();
                            }
                            if (run.FontName is not null)
                            {
                                w.WriteStartElement("a", "latin");
                                w.WriteAttributeString("typeface", run.FontName);
                                w.WriteEndElement();
                            }
                            if (run.Color is not null)
                            {
                                w.WriteStartElement("a", "solidFill");
                                w.WriteStartElement("a", "srgbClr");
                                w.WriteAttributeString("val", run.Color);
                                w.WriteEndElement();
                                w.WriteEndElement();
                            }
                            w.WriteEndElement(); // rPr
                        }
                        w.WriteStartElement("a", "t");
                        if (run.Text is not null)
                            w.WriteString(run.Text);
                        w.WriteEndElement(); // t
                        w.WriteEndElement(); // r
                    }
                    w.WriteEndElement(); // p
                }
                w.WriteEndElement(); // txBody
                w.WriteEndElement(); // tc
            }
            w.WriteEndElement(); // tr
        }

        w.WriteEndElement(); // tbl
        w.WriteEndElement(); // graphicData
        w.WriteEndElement(); // graphic
        w.WriteEndElement(); // p:graphicFrame
    }

    private static void WriteSlideRels(PoiXmlWriter w, XSLFSlide slide)
    {
        w.WriteStartElement("Relationships");
        w.WriteAttributeString("xmlns", "http://schemas.openxmlformats.org/package/2006/relationships");
        WriteRelationship(w, "rId1", "../slideLayouts/slideLayout1.xml", RelTypeSlideLayout);
        foreach (var shape in slide.getShapes())
            WriteRelationship(w, shape.RelationId, $"../media/{shape.PictureData.getFileName()}", RelTypeImage);
        w.WriteEndElement();
    }

    // ----- XML structural helpers -----

    private static void WriteSpTreeStub(PoiXmlWriter w)
    {
        w.WriteStartElement("p", "cSld");
        w.WriteStartElement("p", "spTree");
        w.WriteStartElement("p", "nvGrpSpPr");
        w.WriteStartElement("p", "cNvPr"); w.WriteAttributeString("id", "1"); w.WriteAttributeString("name", ""); w.WriteEndElement();
        w.WriteStartElement("p", "cNvGrpSpPr"); w.WriteEndElement();
        w.WriteStartElement("p", "nvPr"); w.WriteEndElement();
        w.WriteEndElement();
        w.WriteStartElement("p", "grpSpPr");
        w.WriteStartElement("a", "xfrm");
        WriteAOff(w, 0, 0); WriteAExt(w, 0, 0);
        WriteAChOff(w, 0, 0); WriteAChExt(w, 0, 0);
        w.WriteEndElement();
        w.WriteEndElement();
        w.WriteEndElement(); // spTree
        w.WriteEndElement(); // cSld
    }

    private static void WriteAOff(PoiXmlWriter w, long x, long y)
    {
        w.WriteStartElement("a", "off");
        w.WriteAttributeString("x", x.ToString(CultureInfo.InvariantCulture));
        w.WriteAttributeString("y", y.ToString(CultureInfo.InvariantCulture));
        w.WriteEndElement();
    }

    private static void WriteAExt(PoiXmlWriter w, long cx, long cy)
    {
        w.WriteStartElement("a", "ext");
        w.WriteAttributeString("cx", cx.ToString(CultureInfo.InvariantCulture));
        w.WriteAttributeString("cy", cy.ToString(CultureInfo.InvariantCulture));
        w.WriteEndElement();
    }

    private static void WriteAChOff(PoiXmlWriter w, long x, long y)
    {
        w.WriteStartElement("a", "chOff");
        w.WriteAttributeString("x", x.ToString(CultureInfo.InvariantCulture));
        w.WriteAttributeString("y", y.ToString(CultureInfo.InvariantCulture));
        w.WriteEndElement();
    }

    private static void WriteAChExt(PoiXmlWriter w, long cx, long cy)
    {
        w.WriteStartElement("a", "chExt");
        w.WriteAttributeString("cx", cx.ToString(CultureInfo.InvariantCulture));
        w.WriteAttributeString("cy", cy.ToString(CultureInfo.InvariantCulture));
        w.WriteEndElement();
    }

    private static void WriteTxStyleEntry(PoiXmlWriter w, string name)
    {
        w.WriteStartElement("p", name);
        w.WriteStartElement("a", "lstStyle"); w.WriteEndElement();
        w.WriteEndElement();
    }

    private static void WriteDefault(PoiXmlWriter w, string ext, string ct)
    {
        w.WriteStartElement("Default");
        w.WriteAttributeString("Extension", ext);
        w.WriteAttributeString("ContentType", ct);
        w.WriteEndElement();
    }

    private static void WriteOverride(PoiXmlWriter w, string partName, string ct)
    {
        w.WriteStartElement("Override");
        w.WriteAttributeString("PartName", partName);
        w.WriteAttributeString("ContentType", ct);
        w.WriteEndElement();
    }

    private static void WriteRelationship(PoiXmlWriter w, string id, string target, string type)
    {
        w.WriteStartElement("Relationship");
        w.WriteAttributeString("Id", id);
        w.WriteAttributeString("Target", target);
        w.WriteAttributeString("Type", type);
        w.WriteEndElement();
    }

    private static void WriteEntry(ZipArchive archive, string name, Action<PoiXmlWriter> write)
    {
        var entry = archive.CreateEntry(name, CompressionLevel.Optimal);
        using var entryStream = entry.Open();
        using var textWriter  = new StreamWriter(entryStream, new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
        using var writer      = PoiXmlWriterFactory.CreateForOoxmlPackagePart(textWriter, name);
        write(writer);
    }

    private static void WriteBinaryEntry(ZipArchive archive, string name, byte[] data)
    {
        var entry = archive.CreateEntry(name, CompressionLevel.Optimal);
        using var entryStream = entry.Open();
        entryStream.Write(data, 0, data.Length);
    }

    // ----- read (parse) -----

    private void Load(Stream stream)
    {
        _slides.Clear();
        _pictures.Clear();
        _vbaProjectBin = null;

        using var archive = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: true);

        var mediaByName      = LoadMedia(archive);
        var slidePathByRelId = ParsePresentationRels(archive);
        var slideRelIds      = ParseSlideOrder(archive);

        foreach (var slideRelId in slideRelIds)
        {
            if (!slidePathByRelId.TryGetValue(slideRelId, out var slidePath)) continue;
            var fullSlidePath = $"ppt/{slidePath}";
            var slideRelsPath = BuildRelsPath(fullSlidePath);
            var mediaByRid    = ParseSlideRels(archive, slideRelsPath, mediaByName);
            var slide         = new XSLFSlide();
            ParseSlideXml(archive, fullSlidePath, slide, mediaByRid);
            _slides.Add(slide);
        }

        // Read slide size from presentation.xml
        ParseSlideSize(archive);

        // pptm: preserve vbaProject.bin verbatim
        var vba = archive.GetEntry("ppt/vbaProject.bin");
        if (vba != null)
        {
            using var s = vba.Open();
            using var ms = new MemoryStream();
            s.CopyTo(ms);
            if (ms.Length > 0) _vbaProjectBin = ms.ToArray();
        }

        // Collect non-model ZIP entries for unknown-part preservation
        // Must be after _slides is populated so GetModelEntryNames() is accurate
        CollectPreservedEntries(archive);
    }

    private void CollectPreservedEntries(ZipArchive archive)
    {
        _preservedEntries.Clear();
        var known = GetModelEntryNames();
        foreach (var entry in archive.Entries)
        {
            var name = entry.FullName.Replace('\\', '/');
            // Known model entries are NOT preserved (they get re-written)
            if (known.Contains(name, StringComparer.OrdinalIgnoreCase)) continue;
            // Media entries that are already tracked in _pictures (image media) will be re-written; skip them.
            // Entries not in _pictures (video, audio, etc.) are preserved via _preservedEntries
            if (name.StartsWith("ppt/media/", StringComparison.OrdinalIgnoreCase)
                && _pictures.Any(p => string.Equals($"ppt/media/{p.getFileName()}", name, StringComparison.OrdinalIgnoreCase)))
                continue;
            using var s = entry.Open();
            using var ms = new MemoryStream();
            s.CopyTo(ms);
            _preservedEntries[name] = ms.ToArray();
        }
    }

    private HashSet<string> GetModelEntryNames()
    {
        var known = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "[Content_Types].xml",
            "_rels/.rels",
            "ppt/presentation.xml",
            "ppt/_rels/presentation.xml.rels",
            "ppt/presProps.xml",
            "ppt/tableStyles.xml",
            "ppt/theme/theme1.xml",
            "ppt/slideMasters/slideMaster1.xml",
            "ppt/slideMasters/_rels/slideMaster1.xml.rels",
            "ppt/slideLayouts/slideLayout1.xml",
            "ppt/slideLayouts/_rels/slideLayout1.xml.rels",
        };
        // Model-written slides
        for (int i = 1; i <= _slides.Count; i++)
        {
            known.Add($"ppt/slides/slide{i}.xml");
            known.Add($"ppt/slides/_rels/slide{i}.xml.rels");
        }
        return known;
    }

    private static Dictionary<string, byte[]> LoadMedia(ZipArchive archive)
    {
        var result = new Dictionary<string, byte[]>(StringComparer.OrdinalIgnoreCase);
        foreach (var entry in archive.Entries
            .Where(e => e.FullName.StartsWith("ppt/media/", StringComparison.Ordinal)))
        {
            using var s  = entry.Open();
            using var ms = new MemoryStream();
            s.CopyTo(ms);
            result[Path.GetFileName(entry.FullName)] = ms.ToArray();
        }
        return result;
    }

    private static Dictionary<string, string> ParsePresentationRels(ZipArchive archive)
    {
        var map   = new Dictionary<string, string>(StringComparer.Ordinal);
        var entry = archive.GetEntry("ppt/_rels/presentation.xml.rels");
        if (entry is null) return map;
        using var s      = entry.Open();
        using var reader = XmlReader.Create(s);
        while (reader.Read())
        {
            if (reader.NodeType != XmlNodeType.Element || reader.LocalName != "Relationship") continue;
            if (!string.Equals(reader.GetAttribute("Type"), RelTypeSlide, StringComparison.Ordinal)) continue;
            var id     = reader.GetAttribute("Id");
            var target = reader.GetAttribute("Target"); // e.g. "slides/slide1.xml"
            if (id is not null && target is not null) map[id] = target;
        }
        return map;
    }

    private static List<string> ParseSlideOrder(ZipArchive archive)
    {
        var ids   = new List<string>();
        var entry = archive.GetEntry("ppt/presentation.xml");
        if (entry is null) return ids;
        using var s      = entry.Open();
        using var reader = XmlReader.Create(s);
        bool inSldIdLst  = false;
        while (reader.Read())
        {
            if (reader.NodeType == XmlNodeType.Element)
            {
                if (reader.NamespaceURI == NsP && reader.LocalName == "sldIdLst") { inSldIdLst = true; continue; }
                if (inSldIdLst && reader.NamespaceURI == NsP && reader.LocalName == "sldId")
                {
                    var rId = reader.GetAttribute("id", NsR);
                    if (rId is not null) ids.Add(rId);
                }
            }
            else if (reader.NodeType == XmlNodeType.EndElement
                     && reader.NamespaceURI == NsP && reader.LocalName == "sldIdLst")
            {
                inSldIdLst = false;
            }
        }
        return ids;
    }

    private Dictionary<string, XSLFPictureData> ParseSlideRels(
        ZipArchive archive, string relsPath, Dictionary<string, byte[]> mediaByName)
    {
        var map   = new Dictionary<string, XSLFPictureData>(StringComparer.Ordinal);
        var entry = archive.GetEntry(relsPath);
        if (entry is null) return map;
        using var s      = entry.Open();
        using var reader = XmlReader.Create(s);
        while (reader.Read())
        {
            if (reader.NodeType != XmlNodeType.Element || reader.LocalName != "Relationship") continue;
            if (!string.Equals(reader.GetAttribute("Type"), RelTypeImage, StringComparison.Ordinal)) continue;
            var id     = reader.GetAttribute("Id");
            var target = reader.GetAttribute("Target"); // e.g. "../media/image1.jpeg"
            if (id is null || target is null) continue;
            var filename = Path.GetFileName(target);
            if (!mediaByName.TryGetValue(filename, out var bytes)) continue;
            var ext = Path.GetExtension(filename).TrimStart('.').ToLowerInvariant();
            int format;
            try { format = XSLFPictureData.FormatFromExtension(ext); }
            catch (NotSupportedException) { continue; }

            var existing = _pictures.FirstOrDefault(p => p.getFileName() == filename);
            XSLFPictureData pic;
            if (existing is not null)
            {
                pic = existing;
            }
            else
            {
                pic = new XSLFPictureData(bytes, format, _pictures.Count + 1);
                _pictures.Add(pic);
            }
            map[id] = pic;
        }
        return map;
    }

    private static void ParseSlideXml(ZipArchive archive, string slidePath,
        XSLFSlide slide, Dictionary<string, XSLFPictureData> mediaByRid)
    {
        var entry = archive.GetEntry(slidePath);
        if (entry is null) return;
        using var s      = entry.Open();
        using var reader = XmlReader.Create(s);

        // Picture state
        bool inPic = false;
        string? blipEmbed = null;
        long x = 0, y = 0, cx = 0, cy = 0;
        int rot = 0;
        bool flipH = false, flipV = false;
        int shapeId = 0;

        // Auto shape (text box) state
        bool inSp = false;
        bool inTxBody = false;
        bool inAP = false;
        bool inAR = false;
        bool inARPr = false;
        XSLFAutoShape? currentAutoShape = null;
        XSLFTextParagraph? currentParagraph = null;
        XSLFTextRun? currentRun = null;

        // Table (p:graphicFrame) state
        bool inGraphicFrame = false;
        bool inTbl = false;
        bool tableInAP = false;
        bool tableInAR = false;
        bool tableInARPr = false;
        XSLFTable? currentTable = null;
        XSLFTableRow? currentRow = null;
        XSLFTableCell? currentCell = null;
        XSLFTextParagraph? tableCurrentParagraph = null;
        XSLFTextRun? tableCurrentRun = null;
        long tableX = 0, tableY = 0, tableCx = 0, tableCy = 0;

        while (reader.Read())
        {
            if (reader.NodeType == XmlNodeType.Element)
            {
                // --- picture elements ---
                if (reader.NamespaceURI == NsP && reader.LocalName == "pic")
                {
                    inPic = true; blipEmbed = null; x = y = cx = cy = 0; rot = 0; flipH = flipV = false; shapeId = 0;
                }
                if (inPic)
                {
                    if (reader.NamespaceURI == NsP && reader.LocalName == "cNvPr")
                        if (int.TryParse(reader.GetAttribute("id"), out var sid)) shapeId = sid;
                    if (reader.NamespaceURI == NsA && reader.LocalName == "blip")
                        blipEmbed = reader.GetAttribute("embed", NsR);
                    if (reader.NamespaceURI == NsA && reader.LocalName == "xfrm")
                    {
                        if (int.TryParse(reader.GetAttribute("rot"), out var r)) rot = r;
                        flipH = reader.GetAttribute("flipH") == "1";
                        flipV = reader.GetAttribute("flipV") == "1";
                    }
                    if (reader.NamespaceURI == NsA && reader.LocalName == "off")
                    {
                        if (long.TryParse(reader.GetAttribute("x"), out var ox)) x = ox;
                        if (long.TryParse(reader.GetAttribute("y"), out var oy)) y = oy;
                    }
                    if (reader.NamespaceURI == NsA && reader.LocalName == "ext" && cx == 0 && cy == 0)
                    {
                        if (long.TryParse(reader.GetAttribute("cx"), out var ecx)) cx = ecx;
                        if (long.TryParse(reader.GetAttribute("cy"), out var ecy)) cy = ecy;
                    }
                    continue;
                }

                // --- auto shape (text box) elements ---
                if (reader.NamespaceURI == NsP && reader.LocalName == "sp")
                {
                    inSp = true; inTxBody = false; inAP = false; inAR = false;
                    currentAutoShape = new XSLFAutoShape(0);
                    x = y = cx = cy = 0; rot = 0; flipH = flipV = false; shapeId = 0;
                    currentParagraph = null; currentRun = null;
                    continue;
                }
                if (inSp)
                {
                    if (reader.NamespaceURI == NsP && reader.LocalName == "cNvPr" && !inTxBody)
                    {
                        if (int.TryParse(reader.GetAttribute("id"), out var sid)) shapeId = sid;
                        continue;
                    }
                    if (reader.NamespaceURI == NsA && reader.LocalName == "xfrm" && !inTxBody)
                    {
                        if (int.TryParse(reader.GetAttribute("rot"), out var r)) rot = r;
                        flipH = reader.GetAttribute("flipH") == "1";
                        flipV = reader.GetAttribute("flipV") == "1";
                        continue;
                    }
                    if (reader.NamespaceURI == NsA && reader.LocalName == "off" && !inTxBody)
                    {
                        if (long.TryParse(reader.GetAttribute("x"), out var ox)) x = ox;
                        if (long.TryParse(reader.GetAttribute("y"), out var oy)) y = oy;
                        continue;
                    }
                    if (reader.NamespaceURI == NsA && reader.LocalName == "ext" && !inTxBody && cx == 0 && cy == 0)
                    {
                        if (long.TryParse(reader.GetAttribute("cx"), out var ecx)) cx = ecx;
                        if (long.TryParse(reader.GetAttribute("cy"), out var ecy)) cy = ecy;
                        continue;
                    }
                    if (reader.NamespaceURI == NsP && reader.LocalName == "txBody")
                    {
                        inTxBody = true;
                        continue;
                    }
                    if (inTxBody && reader.NamespaceURI == NsA && reader.LocalName == "p")
                    {
                        inAP = true;
                        currentParagraph = new XSLFTextParagraph();
                        continue;
                    }
                    if (inAP && reader.NamespaceURI == NsA && reader.LocalName == "r")
                    {
                        inAR = true;
                        currentRun = new XSLFTextRun();
                        continue;
                    }
                    if (inAR && reader.NamespaceURI == NsA && reader.LocalName == "t")
                    {
                        var text = reader.ReadString();
                        if (currentRun is not null) currentRun.Text = text;
                        continue;
                    }
                    // rPr element inside a:r
                    if (inAR && reader.NamespaceURI == NsA && reader.LocalName == "rPr")
                    {
                        inARPr = true;
                        continue;
                    }
                    // rPr child elements
                    if (inARPr)
                    {
                        if (reader.NamespaceURI == NsA && reader.LocalName == "b")
                            currentRun!.Bold = true;
                        else if (reader.NamespaceURI == NsA && reader.LocalName == "i")
                            currentRun!.Italic = true;
                        else if (reader.NamespaceURI == NsA && reader.LocalName == "u")
                            currentRun!.Underline = true;
                        else if (reader.NamespaceURI == NsA && reader.LocalName == "strike")
                            currentRun!.Strikethrough = true;
                        else if (reader.NamespaceURI == NsA && reader.LocalName == "sz")
                        {
                            var val = reader.GetAttribute("val");
                            if (val is not null && double.TryParse(val, NumberStyles.Integer, CultureInfo.InvariantCulture, out var sz))
                                currentRun!.FontSize = sz / 100.0;
                        }
                        else if (reader.NamespaceURI == NsA && reader.LocalName == "latin")
                        {
                            var tf = reader.GetAttribute("typeface");
                            if (tf is not null) currentRun!.FontName = tf;
                        }
                        else if (reader.NamespaceURI == NsA && reader.LocalName == "srgbClr")
                        {
                            var cv = reader.GetAttribute("val");
                            if (cv is not null) currentRun!.Color = cv;
                        }
                        continue;
                    }
                }

                // --- table (p:graphicFrame) elements ---
                if (reader.NamespaceURI == NsP && reader.LocalName == "graphicFrame")
                {
                    inGraphicFrame = true;
                    inTbl = false; tableInAP = false; tableInAR = false; tableInARPr = false;
                    currentTable = new XSLFTable(0);
                    currentRow = null; currentCell = null;
                    tableCurrentParagraph = null; tableCurrentRun = null;
                    tableX = tableY = tableCx = tableCy = 0;
                    continue;
                }
                if (inGraphicFrame)
                {
                    if (reader.NamespaceURI == NsP && reader.LocalName == "cNvPr" && !inTbl)
                    {
                        if (int.TryParse(reader.GetAttribute("id"), out var sid))
                            currentTable = new XSLFTable(sid);
                        continue;
                    }
                    if (reader.NamespaceURI == NsA && reader.LocalName == "off")
                    {
                        if (long.TryParse(reader.GetAttribute("x"), out var ox)) tableX = ox;
                        if (long.TryParse(reader.GetAttribute("y"), out var oy)) tableY = oy;
                        continue;
                    }
                    if (reader.NamespaceURI == NsA && reader.LocalName == "ext" && !inTbl)
                    {
                        if (long.TryParse(reader.GetAttribute("cx"), out var ecx)) tableCx = ecx;
                        if (long.TryParse(reader.GetAttribute("cy"), out var ecy)) tableCy = ecy;
                        continue;
                    }
                    if (reader.NamespaceURI == NsA && reader.LocalName == "tbl")
                    {
                        inTbl = true;
                        continue;
                    }
                    if (inTbl)
                    {
                        if (reader.NamespaceURI == NsA && reader.LocalName == "gridCol")
                        {
                            var wAttr = reader.GetAttribute("w");
                            if (wAttr is not null && long.TryParse(wAttr, out var w))
                                currentTable?.addGridCol(w);
                            continue;
                        }
                        if (reader.NamespaceURI == NsA && reader.LocalName == "tr")
                        {
                            currentRow = new XSLFTableRow();
                            continue;
                        }
                        if (reader.NamespaceURI == NsA && reader.LocalName == "tc")
                        {
                            currentCell = new XSLFTableCell();
                            tableInAP = false; tableInAR = false; tableInARPr = false;
                            tableCurrentParagraph = null; tableCurrentRun = null;
                            continue;
                        }
                        if (reader.NamespaceURI == NsA && reader.LocalName == "txBody")
                        {
                            // enter cell text body (same structure as auto shape txBody)
                            continue;
                        }
                        // Cell text paragraphs / runs (same as auto shape but with table context)
                        if (reader.NamespaceURI == NsA && reader.LocalName == "p" && currentCell is not null)
                        {
                            tableInAP = true;
                            tableCurrentParagraph = new XSLFTextParagraph();
                            continue;
                        }
                        if (tableInAP && reader.NamespaceURI == NsA && reader.LocalName == "r")
                        {
                            tableInAR = true;
                            tableCurrentRun = new XSLFTextRun();
                            continue;
                        }
                        if (tableInAR && reader.NamespaceURI == NsA && reader.LocalName == "t")
                        {
                            var text = reader.ReadString();
                            if (tableCurrentRun is not null) tableCurrentRun.Text = text;
                            continue;
                        }
                        // rPr element inside a:r (in table cell)
                        if (tableInAR && reader.NamespaceURI == NsA && reader.LocalName == "rPr")
                        {
                            tableInARPr = true;
                            continue;
                        }
                        // rPr child elements (table cell)
                        if (tableInARPr)
                        {
                            if (reader.NamespaceURI == NsA && reader.LocalName == "b")
                                tableCurrentRun!.Bold = true;
                            else if (reader.NamespaceURI == NsA && reader.LocalName == "i")
                                tableCurrentRun!.Italic = true;
                            else if (reader.NamespaceURI == NsA && reader.LocalName == "u")
                                tableCurrentRun!.Underline = true;
                            else if (reader.NamespaceURI == NsA && reader.LocalName == "strike")
                                tableCurrentRun!.Strikethrough = true;
                            else if (reader.NamespaceURI == NsA && reader.LocalName == "sz")
                            {
                                var val = reader.GetAttribute("val");
                                if (val is not null && double.TryParse(val, NumberStyles.Integer, CultureInfo.InvariantCulture, out var sz))
                                    tableCurrentRun!.FontSize = sz / 100.0;
                            }
                            else if (reader.NamespaceURI == NsA && reader.LocalName == "latin")
                            {
                                var tf = reader.GetAttribute("typeface");
                                if (tf is not null) tableCurrentRun!.FontName = tf;
                            }
                            else if (reader.NamespaceURI == NsA && reader.LocalName == "srgbClr")
                            {
                                var cv = reader.GetAttribute("val");
                                if (cv is not null) tableCurrentRun!.Color = cv;
                            }
                            continue;
                        }
                    }
                }
            }
            else if (reader.NodeType == XmlNodeType.EndElement)
            {
                // --- picture end ---
                if (reader.NamespaceURI == NsP && reader.LocalName == "pic" && inPic)
                {
                    inPic = false;
                    if (blipEmbed is not null && mediaByRid.TryGetValue(blipEmbed, out var picData))
                    {
                        var shape = new XSLFPictureShape(picData, blipEmbed, shapeId);
                        shape.setAnchor(x, y, cx, cy);
                        shape.SetRotationAttribute(rot);
                        shape.setFlipHorizontal(flipH);
                        shape.setFlipVertical(flipV);
                        slide.AttachShape(shape);
                    }
                    continue;
                }

                // --- auto shape end ---
                if (reader.NamespaceURI == NsP && reader.LocalName == "sp" && inSp)
                {
                    inSp = false;
                    if (currentAutoShape is not null)
                    {
                        currentAutoShape.setAnchor(x, y, cx, cy);
                        currentAutoShape.SetRotationAttribute(rot);
                        currentAutoShape.setFlipHorizontal(flipH);
                        currentAutoShape.setFlipVertical(flipV);
                        slide.AttachAutoShape(currentAutoShape);
                    }
                    currentAutoShape = null; currentParagraph = null; currentRun = null;
                    continue;
                }
                if (inSp && inTxBody && reader.NamespaceURI == NsP && reader.LocalName == "txBody")
                {
                    inTxBody = false;
                    continue;
                }
                if (inAP && reader.NamespaceURI == NsA && reader.LocalName == "p")
                {
                    inAP = false;
                    if (currentParagraph is not null && currentAutoShape is not null)
                        currentAutoShape.AddParagraph(currentParagraph);
                    currentParagraph = null; currentRun = null;
                    continue;
                }
                if (inAR && reader.NamespaceURI == NsA && reader.LocalName == "r")
                {
                    inAR = false;
                    inARPr = false;
                    if (currentRun is not null && currentParagraph is not null)
                        currentParagraph.AddRun(currentRun);
                    currentRun = null;
                    continue;
                }
                if (inARPr && reader.NamespaceURI == NsA && reader.LocalName == "rPr")
                {
                    inARPr = false;
                    continue;
                }

                // --- table end elements ---
                if (inGraphicFrame && reader.NamespaceURI == NsP && reader.LocalName == "graphicFrame")
                {
                    inGraphicFrame = false;
                    inTbl = false;
                    if (currentTable is not null)
                    {
                        currentTable.setAnchor(tableX, tableY, tableCx, tableCy);
                        slide.AttachTable(currentTable);
                    }
                    currentTable = null; currentRow = null; currentCell = null;
                    continue;
                }
                if (inTbl && reader.NamespaceURI == NsA && reader.LocalName == "tr")
                {
                    if (currentRow is not null && currentTable is not null)
                        currentTable.AddRow(currentRow);
                    currentRow = null;
                    continue;
                }
                if (inTbl && reader.NamespaceURI == NsA && reader.LocalName == "tc")
                {
                    if (currentCell is not null && currentRow is not null)
                        currentRow.AddCell(currentCell);
                    currentCell = null;
                    continue;
                }
                if (inTbl && reader.NamespaceURI == NsA && reader.LocalName == "txBody")
                {
                    // exiting cell text body
                    tableInAP = false; tableInAR = false; tableInARPr = false;
                    tableCurrentParagraph = null; tableCurrentRun = null;
                    continue;
                }
                if (inTbl && tableInAP && reader.NamespaceURI == NsA && reader.LocalName == "p")
                {
                    tableInAP = false;
                    if (tableCurrentParagraph is not null && currentCell is not null)
                        currentCell.AddParagraph(tableCurrentParagraph);
                    tableCurrentParagraph = null; tableCurrentRun = null;
                    continue;
                }
                if (inTbl && tableInAR && reader.NamespaceURI == NsA && reader.LocalName == "r")
                {
                    tableInAR = false;
                    tableInARPr = false;
                    if (tableCurrentRun is not null && tableCurrentParagraph is not null)
                        tableCurrentParagraph.AddRun(tableCurrentRun);
                    tableCurrentRun = null;
                    continue;
                }
                if (inTbl && tableInARPr && reader.NamespaceURI == NsA && reader.LocalName == "rPr")
                {
                    tableInARPr = false;
                    continue;
                }
            }
        }
    }

    private static string BuildRelsPath(string partPath)
    {
        var dir  = Path.GetDirectoryName(partPath)?.Replace('\\', '/') ?? string.Empty;
        var file = Path.GetFileName(partPath);
        return $"{dir}/_rels/{file}.rels";
    }

    private void ParseSlideSize(ZipArchive archive)
    {
        var entry = archive.GetEntry("ppt/presentation.xml");
        if (entry is null) return;
        using var s      = entry.Open();
        using var reader = XmlReader.Create(s);
        while (reader.Read())
        {
            if (reader.NodeType == XmlNodeType.Element && reader.NamespaceURI == NsP && reader.LocalName == "sldSz")
            {
                var cxAttr = reader.GetAttribute("cx");
                var cyAttr = reader.GetAttribute("cy");
                if (cxAttr is not null && long.TryParse(cxAttr, NumberStyles.Integer, CultureInfo.InvariantCulture, out var cx))
                    _slideCx = cx;
                if (cyAttr is not null && long.TryParse(cyAttr, NumberStyles.Integer, CultureInfo.InvariantCulture, out var cy))
                    _slideCy = cy;
                return;
            }
        }
    }
}
