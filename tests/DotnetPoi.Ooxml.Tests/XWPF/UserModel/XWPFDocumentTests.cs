using System.IO.Compression;
using DotnetPoi.XWPF.UserModel;
using Xunit;

namespace DotnetPoi.XWPF.Tests.UserModel;

public class XWPFDocumentTests
{
    private static byte[] LoadTestImage() => File.ReadAllBytes("image.jpg");

    [Fact]
    public void Write_EmptyDocument_ProducesValidDocx()
    {
        using var doc = new XWPFDocument();
        using var stream = new MemoryStream();
        doc.write(stream);

        stream.Position = 0;
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read);
        Assert.NotNull(archive.GetEntry("[Content_Types].xml"));
        Assert.NotNull(archive.GetEntry("_rels/.rels"));
        Assert.NotNull(archive.GetEntry("word/document.xml"));
    }

    [Fact]
    public void Write_AnyDocument_AlwaysIncludesSettingsAndRels()
    {
        using var doc = new XWPFDocument();
        doc.createParagraph().createRun().setText("hello");

        using var stream = new MemoryStream();
        doc.write(stream);

        stream.Position = 0;
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read);
        Assert.NotNull(archive.GetEntry("word/settings.xml"));
        Assert.NotNull(archive.GetEntry("word/_rels/document.xml.rels"));
        var rels = ReadEntry(archive, "word/_rels/document.xml.rels");
        Assert.Contains("rId1", rels);
        Assert.Contains("settings.xml", rels);
        var ct = ReadEntry(archive, "[Content_Types].xml");
        Assert.Contains("wordprocessingml.settings", ct);
    }

    [Fact]
    public void Write_SimpleParagraph_TextAppearsInDocumentXml()
    {
        using var doc = new XWPFDocument();
        var para = doc.createParagraph();
        para.createRun().setText("Hello World");

        using var stream = new MemoryStream();
        doc.write(stream);

        stream.Position = 0;
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read);
        var documentXml = ReadEntry(archive, "word/document.xml");
        Assert.Contains("Hello World", documentXml);
        Assert.Contains("<w:t", documentXml);
    }

    [Fact]
    public void Write_BoldItalicRun_WritesRprElements()
    {
        using var doc = new XWPFDocument();
        var para = doc.createParagraph();
        var run = para.createRun();
        run.setText("styled");
        run.setBold(true);
        run.setItalic(true);

        using var stream = new MemoryStream();
        doc.write(stream);

        stream.Position = 0;
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read);
        var xml = ReadEntry(archive, "word/document.xml");
        // POI writes w:val="on" explicitly — we match this format
        Assert.Contains("<w:b w:val=\"on\"/>", xml);
        Assert.Contains("<w:i w:val=\"on\"/>", xml);
        Assert.Contains("<w:rPr>", xml);
    }

    [Fact]
    public void Write_InlinePicture_ProducesMediaAndRelationship()
    {
        var imageBytes = LoadTestImage();
        using var doc = new XWPFDocument();
        var para = doc.createParagraph();
        var run = para.createRun();
        run.addPicture(imageBytes, XWPFPictureData.PICTURE_TYPE_JPEG, "test.jpeg",
            914400, 914400); // 1 inch × 1 inch

        using var stream = new MemoryStream();
        doc.write(stream);

        stream.Position = 0;
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read);

        Assert.NotNull(archive.GetEntry("word/media/image1.jpeg"));
        Assert.NotNull(archive.GetEntry("word/_rels/document.xml.rels"));

        var rels = ReadEntry(archive, "word/_rels/document.xml.rels");
        Assert.Contains("Target=\"settings.xml\"", rels);
        Assert.Contains("Target=\"media/image1.jpeg\"", rels);
        Assert.Contains("Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\"", rels);

        var xml = ReadEntry(archive, "word/document.xml");
        Assert.Contains("<w:drawing>", xml);
        Assert.Contains("<wp:inline", xml);
        // rId1 = settings.xml; rId2 = styles.xml; image starts at rId3
        Assert.Contains("r:embed=\"rId3\"", xml);
    }

    [Fact]
    public void Write_InlinePicture_MediaBytesStoredFaithfully()
    {
        var imageBytes = LoadTestImage();
        using var doc = new XWPFDocument();
        doc.createParagraph().createRun().addPicture(imageBytes, XWPFPictureData.PICTURE_TYPE_JPEG, "img.jpeg",
            914400, 914400);

        using var stream = new MemoryStream();
        doc.write(stream);

        stream.Position = 0;
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read);
        using var entry = archive.GetEntry("word/media/image1.jpeg")!.Open();
        using var ms = new MemoryStream();
        entry.CopyTo(ms);
        Assert.Equal(imageBytes, ms.ToArray());
    }

    [Fact]
    public void Write_RotatedPicture_WritesRotAttribute()
    {
        var imageBytes = LoadTestImage();
        using var doc = new XWPFDocument();
        var run = doc.createParagraph().createRun();
        var picture = run.addPicture(imageBytes, XWPFPictureData.PICTURE_TYPE_JPEG, "img.jpeg",
            914400, 914400);
        picture.setRotation(90.0);

        using var stream = new MemoryStream();
        doc.write(stream);

        stream.Position = 0;
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read);
        var xml = ReadEntry(archive, "word/document.xml");
        // 90° = 90 × 60000 = 5400000
        Assert.Contains("rot=\"5400000\"", xml);
    }

    [Fact]
    public void Write_TwoPictures_BothMediaFilesPresent()
    {
        var imageBytes = LoadTestImage();
        using var doc = new XWPFDocument();
        var para = doc.createParagraph();
        var run1 = para.createRun();
        run1.addPicture(imageBytes, XWPFPictureData.PICTURE_TYPE_JPEG, "a.jpeg", 100000, 100000);
        var run2 = para.createRun();
        run2.addPicture(imageBytes, XWPFPictureData.PICTURE_TYPE_JPEG, "b.jpeg", 100000, 100000);

        using var stream = new MemoryStream();
        doc.write(stream);

        stream.Position = 0;
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read);
        // Same bytes → deduplicated to one media file
        Assert.NotNull(archive.GetEntry("word/media/image1.jpeg"));
    }

    [Fact]
    public void Rotation_RoundTrip_PreservesValue()
    {
        var imageBytes = LoadTestImage();
        using var original = new XWPFDocument();
        var run = original.createParagraph().createRun();
        var picture = run.addPicture(imageBytes, XWPFPictureData.PICTURE_TYPE_JPEG, "img.jpeg",
            914400, 914400);
        picture.setRotation(45.0);

        using var stream = new MemoryStream();
        original.write(stream);

        stream.Position = 0;
        using var loaded = new XWPFDocument(stream);

        Assert.Single(loaded.getParagraphs());
        Assert.Single(loaded.getParagraphs()[0].getRuns());
        var loadedPicture = Assert.Single(loaded.getParagraphs()[0].getRuns()[0].getEmbeddedPictures());
        Assert.Equal(45.0, loadedPicture.getRotation(), precision: 6);
    }

    [Fact]
    public void SetRotation_NormalisesNegativeAngles()
    {
        var imageBytes = LoadTestImage();
        using var doc = new XWPFDocument();
        var run = doc.createParagraph().createRun();
        var picture = run.addPicture(imageBytes, XWPFPictureData.PICTURE_TYPE_JPEG, "img.jpeg", 100, 100);

        picture.setRotation(-90.0);
        Assert.Equal(270.0, picture.getRotation(), precision: 6);
    }

    [Fact]
    public void Read_WrittenDocument_RestoresTextAndPicture()
    {
        var imageBytes = LoadTestImage();
        using var original = new XWPFDocument();
        var para = original.createParagraph();
        var textRun = para.createRun();
        textRun.setText("test text");
        textRun.setBold(true);
        var imageRun = para.createRun();
        imageRun.addPicture(imageBytes, XWPFPictureData.PICTURE_TYPE_JPEG, "img.jpeg", 914400, 914400);

        using var stream = new MemoryStream();
        original.write(stream);

        stream.Position = 0;
        using var loaded = new XWPFDocument(stream);

        Assert.Single(loaded.getParagraphs());
        var loadedPara = loaded.getParagraphs()[0];
        Assert.Equal("test text", loadedPara.getText());
        // Bold attribute on the text run must survive round-trip.
        Assert.True(loadedPara.getRuns()[0].isBold());
        Assert.Single(loaded.getAllPictures());
        Assert.Equal(XWPFPictureData.PICTURE_TYPE_JPEG, loaded.getAllPictures()[0].getPictureType());
        Assert.Equal(imageBytes, loaded.getAllPictures()[0].getData());
    }

    [Fact]
    public void RoundTrip_MultipleParagraphs_TextAndFormattingRestored()
    {
        using var original = new XWPFDocument();

        var p1 = original.createParagraph();
        var r1 = p1.createRun();
        r1.setText("bold text");
        r1.setBold(true);

        var p2 = original.createParagraph();
        var r2 = p2.createRun();
        r2.setText("italic text");
        r2.setItalic(true);

        var p3 = original.createParagraph();
        var r3a = p3.createRun();
        r3a.setText("plain ");
        var r3b = p3.createRun();
        r3b.setText("bold");
        r3b.setBold(true);

        using var stream = new MemoryStream();
        original.write(stream);

        stream.Position = 0;
        using var loaded = new XWPFDocument(stream);

        Assert.Equal(3, loaded.getParagraphs().Count);

        var lp1 = loaded.getParagraphs()[0];
        Assert.Equal("bold text", lp1.getText());
        Assert.True(lp1.getRuns()[0].isBold());
        Assert.False(lp1.getRuns()[0].isItalic());

        var lp2 = loaded.getParagraphs()[1];
        Assert.Equal("italic text", lp2.getText());
        Assert.False(lp2.getRuns()[0].isBold());
        Assert.True(lp2.getRuns()[0].isItalic());

        var lp3 = loaded.getParagraphs()[2];
        Assert.Equal("plain bold", lp3.getText());
        Assert.Equal(2, lp3.getRuns().Count);
        Assert.Equal("plain ", lp3.getRuns()[0].getText(0));
        Assert.False(lp3.getRuns()[0].isBold());
        Assert.Equal("bold", lp3.getRuns()[1].getText(0));
        Assert.True(lp3.getRuns()[1].isBold());
    }

    [Fact]
    public void RoundTrip_FontProperties_Restored()
    {
        using var original = new XWPFDocument();
        var para = original.createParagraph();
        var run = para.createRun();
        run.setText("styled text");
        run.setBold(true);
        run.setItalic(true);
        run.setFontName("Arial");
        run.setFontSize(12.0);
        run.setColor("FF0000");
        run.setUnderline(true);
        run.setStrike(true);
        para.setAlignment(ParagraphAlignment.Center);

        using var stream = new MemoryStream();
        original.write(stream);
        stream.Position = 0;

        using var loaded = new XWPFDocument(stream);
        Assert.Single(loaded.getParagraphs());
        var loadedPara = loaded.getParagraphs()[0];
        Assert.Single(loadedPara.getRuns());
        var loadedRun = loadedPara.getRuns()[0];
        Assert.Equal("styled text", loadedRun.getText(0));
        Assert.True(loadedRun.isBold());
        Assert.True(loadedRun.isItalic());
        Assert.Equal("Arial", loadedRun.getFontName());
        Assert.Equal(12.0, loadedRun.getFontSize(), precision: 1);
        Assert.Equal("FF0000", loadedRun.getColor());
        Assert.True(loadedRun.isUnderline());
        Assert.True(loadedRun.isStrike());
        Assert.Equal(ParagraphAlignment.Center, loadedPara.getAlignment());
    }

    [Fact]
    public void RoundTrip_RunFontNameSizeColor_Restored()
    {
        using var original = new XWPFDocument();
        var run = original.createParagraph().createRun();
        run.setText("quick brown fox");
        run.setFontName("Times New Roman");
        run.setFontSize(14.0);
        run.setColor("00FF00");
        run.setBold(false); // no rPr at all
        run.setItalic(false);

        using var stream = new MemoryStream();
        original.write(stream);
        stream.Position = 0;

        using var loaded = new XWPFDocument(stream);
        var loadedRun = loaded.getParagraphs()[0].getRuns()[0];
        Assert.Equal("quick brown fox", loadedRun.getText(0));
        Assert.Equal("Times New Roman", loadedRun.getFontName());
        Assert.Equal(14.0, loadedRun.getFontSize(), precision: 1);
        Assert.Equal("00FF00", loadedRun.getColor());
        Assert.False(loadedRun.isBold());
        Assert.False(loadedRun.isItalic());
    }

    [Fact]
    public void RoundTrip_AlignmentAllValues_Restored()
    {
        foreach (var align in new[] { ParagraphAlignment.Left, ParagraphAlignment.Center,
            ParagraphAlignment.Right, ParagraphAlignment.Both })
        {
            using var original = new XWPFDocument();
            var para = original.createParagraph();
            para.createRun().setText("text");
            para.setAlignment(align);

            using var stream = new MemoryStream();
            original.write(stream);
            stream.Position = 0;

            using var loaded = new XWPFDocument(stream);
            Assert.Equal(align, loaded.getParagraphs()[0].getAlignment());
        }
    }

    [Fact]
    public void RoundTrip_ParagraphIndentation_Restored()
    {
        using var original = new XWPFDocument();
        var para = original.createParagraph();
        para.createRun().setText("indented");
        para.setIndentationLeft(720);   // 0.5 inch
        para.setIndentationRight(360);  // 0.25 inch
        para.setIndentationFirstLine(720);

        using var stream = new MemoryStream();
        original.write(stream);
        stream.Position = 0;

        using var loaded = new XWPFDocument(stream);
        var loadedPara = loaded.getParagraphs()[0];
        Assert.Equal(720, loadedPara.getIndentationLeft());
        Assert.Equal(360, loadedPara.getIndentationRight());
        Assert.Equal(720, loadedPara.getIndentationFirstLine());
    }

    [Fact]
    public void RoundTrip_ParagraphSpacing_Restored()
    {
        using var original = new XWPFDocument();
        var para = original.createParagraph();
        para.createRun().setText("spaced");
        para.setSpacingBefore(240);
        para.setSpacingAfter(120);
        para.setSpacingBetween(400);
        para.setLineSpacingRule(LineSpacingRule.Exact);

        using var stream = new MemoryStream();
        original.write(stream);
        stream.Position = 0;

        using var loaded = new XWPFDocument(stream);
        var loadedPara = loaded.getParagraphs()[0];
        Assert.Equal(240, loadedPara.getSpacingBefore());
        Assert.Equal(120, loadedPara.getSpacingAfter());
        Assert.Equal(400, loadedPara.getSpacingBetween());
        Assert.Equal(LineSpacingRule.Exact, loadedPara.getLineSpacingRule());
    }

    [Fact]
    public void RoundTrip_BulletList_Restored()
    {
        using var original = new XWPFDocument();
        var p1 = original.createParagraph();
        p1.createRun().setText("item 1");
        p1.setBulletList();
        var p2 = original.createParagraph();
        p2.createRun().setText("item 2");
        p2.setBulletList();

        using var stream = new MemoryStream();
        original.write(stream);
        stream.Position = 0;

        using var loaded = new XWPFDocument(stream);
        Assert.Equal(2, loaded.getParagraphs().Count);
        Assert.NotNull(loaded.getParagraphs()[0].getNumId());
        Assert.Equal(0, loaded.getParagraphs()[0].getIlvl());
        Assert.NotNull(loaded.getParagraphs()[1].getNumId());
        Assert.Equal(0, loaded.getParagraphs()[1].getIlvl());
        Assert.Equal("item 1", loaded.getParagraphs()[0].getText());
        Assert.Equal("item 2", loaded.getParagraphs()[1].getText());
    }

    [Fact]
    public void RoundTrip_NumberedList_Restored()
    {
        using var original = new XWPFDocument();
        var para = original.createParagraph();
        para.createRun().setText("numbered item");
        para.setNumberedList();

        using var stream = new MemoryStream();
        original.write(stream);
        stream.Position = 0;

        using var loaded = new XWPFDocument(stream);
        var loadedPara = loaded.getParagraphs()[0];
        Assert.NotNull(loadedPara.getNumId());
        Assert.Equal(0, loadedPara.getIlvl());
        Assert.Equal("numbered item", loadedPara.getText());
    }

    [Fact]
    public void RoundTrip_Table_Restored()
    {
        using var original = new XWPFDocument();
        var table = original.createTable();
        table.addGridCol(9144); // ~1 inch
        table.addGridCol(9144);
        var row = table.createRow();
        var cell1 = row.createCell();
        cell1.addParagraph().createRun().setText("A1");
        var cell2 = row.createCell();
        cell2.addParagraph().createRun().setText("B1");
        var row2 = table.createRow();
        row2.createCell().addParagraph().createRun().setText("A2");
        row2.createCell().addParagraph().createRun().setText("B2");

        using var stream = new MemoryStream();
        original.write(stream);
        stream.Position = 0;

        using var loaded = new XWPFDocument(stream);
        var tables = loaded.getTables();
        Assert.Single(tables);
        var loadedTable = tables[0];
        Assert.Equal(2, loadedTable.getRows().Count);
        Assert.Equal(2, loadedTable.getRows()[0].getCells().Count);
        Assert.Equal(2, loadedTable.getRows()[1].getCells().Count);
        Assert.Equal("A1", loadedTable.getRows()[0].getCells()[0].getParagraphs()[0].getText());
        Assert.Equal("B2", loadedTable.getRows()[1].getCells()[1].getParagraphs()[0].getText());
    }

    [Fact]
    public void RoundTrip_TableAndParagraphs_Combined()
    {
        using var original = new XWPFDocument();
        var p1 = original.createParagraph();
        p1.createRun().setText("before table");

        var table = original.createTable();
        table.addGridCol(9144);
        var row = table.createRow();
        row.createCell().addParagraph().createRun().setText("cell");

        var p2 = original.createParagraph();
        p2.createRun().setText("after table");

        using var stream = new MemoryStream();
        original.write(stream);
        stream.Position = 0;

        using var loaded = new XWPFDocument(stream);
        Assert.Equal(2, loaded.getParagraphs().Count);
        Assert.Single(loaded.getTables());
        Assert.Equal("before table", loaded.getParagraphs()[0].getText());
        Assert.Equal("after table", loaded.getParagraphs()[1].getText());
        Assert.Equal("cell", loaded.getTables()[0].getRows()[0].getCells()[0].getParagraphs()[0].getText());
    }

    [Fact]
    public void RoundTrip_Hyperlink_Restored()
    {
        using var original = new XWPFDocument();
        var para = original.createParagraph();
        var run = para.createRun();
        run.setText("Click here");
        run.setHyperlink("https://example.com");

        using var stream = new MemoryStream();
        original.write(stream);
        stream.Position = 0;

        using var loaded = new XWPFDocument(stream);
        var loadedPara = loaded.getParagraphs()[0];
        var loadedRun = loadedPara.getRuns()[0];
        Assert.Equal("Click here", loadedRun.getText(0));
        Assert.Equal("https://example.com", loadedRun.getHyperlink());
    }

    [Fact]
    public void RoundTrip_PageSetup_Restored()
    {
        using var original = new XWPFDocument();
        original.setPageSize(15840, 12240); // landscape-ish dimensions
        original.setLandscape(true);
        original.setMargins(720, 720, 720, 720); // 0.5in all sides
        original.createParagraph().createRun().setText("page setup test");

        using var stream = new MemoryStream();
        original.write(stream);
        stream.Position = 0;

        using var loaded = new XWPFDocument(stream);
        Assert.Equal(15840, loaded.getPageWidth());
        Assert.Equal(12240, loaded.getPageHeight());
        Assert.True(loaded.isLandscape());
        Assert.Equal(720, loaded.getMarginTop());
        Assert.Equal(720, loaded.getMarginRight());
        Assert.Equal(720, loaded.getMarginBottom());
        Assert.Equal(720, loaded.getMarginLeft());
    }

    [Fact]
    public void RoundTrip_HeaderFooter_Restored()
    {
        using var original = new XWPFDocument();
        original.setHeaderText("My Header");
        original.setFooterText("Page 1");
        original.createParagraph().createRun().setText("header/footer test");

        using var stream = new MemoryStream();
        original.write(stream);
        stream.Position = 0;

        using var loaded = new XWPFDocument(stream);
        Assert.Equal("My Header", loaded.getHeaderText());
        Assert.Equal("Page 1", loaded.getFooterText());
    }

    [Fact]
    public void RoundTrip_HeaderFooterVariants_Restored()
    {
        using var original = new XWPFDocument();
        original.setHeaderText("Default Header");
        original.setFirstHeaderText("First Page Header");
        original.setEvenHeaderText("Even Page Header");
        original.setFooterText("Default Footer");
        original.setFirstFooterText("First Page Footer");
        original.setEvenFooterText("Even Page Footer");
        original.createParagraph().createRun().setText("header/footer variant test");

        using var stream = new MemoryStream();
        original.write(stream);
        stream.Position = 0;

        // Verify XML contains all three headerReference types
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read);
        var docEntry = archive.GetEntry("word/document.xml");
        Assert.NotNull(docEntry);
        using var r = new StreamReader(docEntry.Open());
        var docXml = r.ReadToEnd();
        Assert.Contains("w:type=\"default\"", docXml);
        Assert.Contains("w:type=\"first\"", docXml);
        Assert.Contains("w:type=\"even\"", docXml);

        // Verify rels contain all three headers
        var relsEntry = archive.GetEntry("word/_rels/document.xml.rels");
        Assert.NotNull(relsEntry);
        using var relsReader = new StreamReader(relsEntry.Open());
        var relsXml = relsReader.ReadToEnd();
        Assert.Contains("header1.xml", relsXml);
        Assert.Contains("header2.xml", relsXml);
        Assert.Contains("header3.xml", relsXml);
        Assert.Contains("footer1.xml", relsXml);
        Assert.Contains("footer2.xml", relsXml);
        Assert.Contains("footer3.xml", relsXml);
        Assert.Contains("Id=\"rId3\" Target=\"header1.xml\"", relsXml);
        Assert.Contains("Id=\"rId6\" Target=\"footer1.xml\"", relsXml);
        Assert.Contains("r:id=\"rId3\"", docXml);
        Assert.Contains("r:id=\"rId6\"", docXml);

        // Verify content types
        var ctEntry = archive.GetEntry("[Content_Types].xml");
        Assert.NotNull(ctEntry);
        using var ctReader = new StreamReader(ctEntry.Open());
        var ctXml = ctReader.ReadToEnd();
        Assert.Contains("/word/header1.xml", ctXml);
        Assert.Contains("/word/header2.xml", ctXml);
        Assert.Contains("/word/header3.xml", ctXml);
        Assert.Contains("/word/footer1.xml", ctXml);
        Assert.Contains("/word/footer2.xml", ctXml);
        Assert.Contains("/word/footer3.xml", ctXml);

        // Load and verify API
        stream.Position = 0;
        using var loaded = new XWPFDocument(stream);
        Assert.Equal("Default Header", loaded.getHeaderText());
        Assert.Equal("First Page Header", loaded.getFirstHeaderText());
        Assert.Equal("Even Page Header", loaded.getEvenHeaderText());
        Assert.Equal("Default Footer", loaded.getFooterText());
        Assert.Equal("First Page Footer", loaded.getFirstFooterText());
        Assert.Equal("Even Page Footer", loaded.getEvenFooterText());
    }

    [Fact]
    public void RoundTrip_UnknownParts_Preserved()
    {
        using var original = new XWPFDocument();
        original.createParagraph().createRun().setText("test");

        byte[] raw;
        using (var ms = new MemoryStream())
        {
            original.write(ms);
            raw = ms.ToArray();
        }

        // Re-pack with extra non-model entries
        using var injectedStream = new MemoryStream();
        using (var writerArchive = new ZipArchive(injectedStream, ZipArchiveMode.Create, leaveOpen: true))
        {
            using (var readerArchive = new ZipArchive(new MemoryStream(raw), ZipArchiveMode.Read))
            {
                foreach (var entry in readerArchive.Entries)
                {
                    using var s = entry.Open();
                    using var ms = new MemoryStream();
                    s.CopyTo(ms);
                    var newEntry = writerArchive.CreateEntry(entry.FullName);
                    using var ws = newEntry.Open();
                    ws.Write(ms.ToArray(), 0, (int)ms.Length);
                }
            }
            var extraEntry = writerArchive.CreateEntry("word/styles.xml");
            using (var ws = extraEntry.Open())
                ws.Write(System.Text.Encoding.UTF8.GetBytes("<w:styles/>"));
            extraEntry = writerArchive.CreateEntry("docProps/custom.xml");
            using (var ws = extraEntry.Open())
                ws.Write(System.Text.Encoding.UTF8.GetBytes("<Properties/>"));
        }
        injectedStream.Position = 0;

        using var loaded = new XWPFDocument(injectedStream);
        Assert.Equal("test", loaded.getParagraphs()[0].getText());

        // Write again and verify extra entries survived
        using var outStream = new MemoryStream();
        loaded.write(outStream);
        outStream.Position = 0;

        using var verifyArchive = new ZipArchive(outStream, ZipArchiveMode.Read);
        var entryNames = verifyArchive.Entries.Select(e => e.FullName).ToHashSet(StringComparer.OrdinalIgnoreCase);
        Assert.Contains("word/styles.xml", entryNames);
        Assert.Contains("docProps/custom.xml", entryNames);
        Assert.Contains("word/document.xml", entryNames);
    }

    [Fact]
    public void RoundTrip_Field_Preserved()
    {
        using var original = new XWPFDocument();
        var para = original.createParagraph();
        para.createRun().setText("Before field. ");
        para.addField(" PAGE ", "1");
        para.createRun().setText(" After field.");

        using var stream = new MemoryStream();
        original.write(stream);
        stream.Position = 0;

        // Verify raw XML contains the field structure
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read);
        var docXml = ReadEntry(archive, "word/document.xml");
        Assert.Contains("fldChar", docXml);
        Assert.Contains("instrText", docXml);
        Assert.Contains("PAGE", docXml);

        // Reload and verify fields are restored
        stream.Position = 0;
        using var loaded = new XWPFDocument(stream);
        var loadedPara = loaded.getParagraphs()[0];
        var fields = loadedPara.getFields();
        Assert.Single(fields);
        Assert.Equal(" PAGE ", fields[0].Instruction);
        Assert.Equal("1", fields[0].Result);

        Assert.Equal("Before field. 1 After field.", loadedPara.getText());
    }

    [Fact]
    public void RoundTrip_TrackedChanges_PreservedInParagraphOrder()
    {
        using var stream = CreateDocxWithDocumentXml("""
            <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
              <w:body>
                <w:p>
                  <w:r><w:t>Before </w:t></w:r>
                  <w:ins w:id="1" w:author="Alice" w:date="2026-05-08T00:00:00Z">
                    <w:r><w:t>inserted </w:t></w:r>
                  </w:ins>
                  <w:r><w:t>Middle </w:t></w:r>
                  <w:del w:id="2" w:author="Bob" w:date="2026-05-08T00:00:00Z">
                    <w:r><w:delText>deleted </w:delText></w:r>
                  </w:del>
                  <w:r><w:t>After</w:t></w:r>
                </w:p>
              </w:body>
            </w:document>
            """);

        using var doc = new XWPFDocument(stream);
        using var outStream = new MemoryStream();
        doc.write(outStream);
        outStream.Position = 0;

        using var archive = new ZipArchive(outStream, ZipArchiveMode.Read);
        var xml = ReadEntry(archive, "word/document.xml");

        var before = xml.IndexOf("Before ", StringComparison.Ordinal);
        var ins = xml.IndexOf("<w:ins", StringComparison.Ordinal);
        var middle = xml.IndexOf("Middle ", StringComparison.Ordinal);
        var del = xml.IndexOf("<w:del", StringComparison.Ordinal);
        var after = xml.IndexOf("After", StringComparison.Ordinal);

        Assert.True(before >= 0);
        Assert.True(ins > before);
        Assert.True(middle > ins);
        Assert.True(del > middle);
        Assert.True(after > del);
        Assert.Contains("w:author=\"Alice\"", xml);
        Assert.Contains("<w:delText>deleted </w:delText>", xml);
    }

    [Fact]
    public void RoundTrip_TrackedChanges_PreservedInTableCellParagraphOrder()
    {
        using var stream = CreateDocxWithDocumentXml("""
            <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
              <w:body>
                <w:tbl>
                  <w:tr>
                    <w:tc>
                      <w:p>
                        <w:r><w:t>Cell before </w:t></w:r>
                        <w:ins w:id="3" w:author="Carol"><w:r><w:t>cell insert </w:t></w:r></w:ins>
                        <w:r><w:t>Cell after</w:t></w:r>
                      </w:p>
                    </w:tc>
                  </w:tr>
                </w:tbl>
              </w:body>
            </w:document>
            """);

        using var doc = new XWPFDocument(stream);
        using var outStream = new MemoryStream();
        doc.write(outStream);
        outStream.Position = 0;

        using var archive = new ZipArchive(outStream, ZipArchiveMode.Read);
        var xml = ReadEntry(archive, "word/document.xml");

        var before = xml.IndexOf("Cell before ", StringComparison.Ordinal);
        var ins = xml.IndexOf("<w:ins", StringComparison.Ordinal);
        var after = xml.IndexOf("Cell after", StringComparison.Ordinal);

        Assert.True(before >= 0);
        Assert.True(ins > before);
        Assert.True(after > ins);
        Assert.Contains("w:author=\"Carol\"", xml);
    }

    [Fact]
    public void RoundTrip_PoiDocxCommentMarkersAndParts_Preserved()
    {
        using var input = new MemoryStream(File.ReadAllBytes(GetPoiTestDataPath("document/testComment.docx")));
        using var doc = new XWPFDocument(input);

        using var outStream = new MemoryStream();
        doc.write(outStream);
        outStream.Position = 0;

        using var archive = new ZipArchive(outStream, ZipArchiveMode.Read);
        var xml = ReadEntry(archive, "word/document.xml");
        AssertOrder(xml, "this is a ", "<w:commentRangeStart", "comment ", "<w:commentRangeEnd", "<w:commentReference", "paragraph!");

        Assert.NotNull(archive.GetEntry("word/comments.xml"));
        Assert.NotNull(archive.GetEntry("word/commentsExtended.xml"));
        Assert.NotNull(archive.GetEntry("word/commentsIds.xml"));
        Assert.NotNull(archive.GetEntry("word/commentsExtensible.xml"));

        var rels = ReadEntry(archive, "word/_rels/document.xml.rels");
        Assert.Contains("relationships/comments", rels);
        Assert.Contains("comments.xml", rels);
        Assert.Contains("commentsExtended.xml", rels);
        Assert.Contains("commentsIds.xml", rels);
        Assert.Contains("commentsExtensible.xml", rels);

        var contentTypes = ReadEntry(archive, "[Content_Types].xml");
        Assert.Contains("/word/comments.xml", contentTypes);
        Assert.Contains("wordprocessingml.comments+xml", contentTypes);
        Assert.Contains("commentsExtended+xml", contentTypes);
        Assert.Contains("commentsIds+xml", contentTypes);
        Assert.Contains("commentsExtensible+xml", contentTypes);
    }

    [Fact]
    public void RoundTrip_CollapsedDocxCommentReferenceInsideRun_Preserved()
    {
        using var input = new MemoryStream(File.ReadAllBytes(GetPoiTestDataPath("document/comment.docx")));
        using var doc = new XWPFDocument(input);

        using var outStream = new MemoryStream();
        doc.write(outStream);
        outStream.Position = 0;

        using var archive = new ZipArchive(outStream, ZipArchiveMode.Read);
        var xml = ReadEntry(archive, "word/document.xml");

        Assert.Contains("<w:commentReference", xml);
        Assert.Contains("w:id=\"0\"", xml);
        Assert.NotNull(archive.GetEntry("word/comments.xml"));
    }

    [Fact]
    public void RoundTrip_DocxCommentMarkers_PreservedInTableCellParagraphOrder()
    {
        using var stream = CreateDocxWithDocumentXml("""
            <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
              <w:body>
                <w:tbl>
                  <w:tr>
                    <w:tc>
                      <w:p>
                        <w:r><w:t>Cell before </w:t></w:r>
                        <w:commentRangeStart w:id="7"/>
                        <w:r><w:t>commented cell </w:t></w:r>
                        <w:commentRangeEnd w:id="7"/>
                        <w:r><w:commentReference w:id="7"/></w:r>
                        <w:r><w:t>Cell after</w:t></w:r>
                      </w:p>
                    </w:tc>
                  </w:tr>
                </w:tbl>
              </w:body>
            </w:document>
            """);

        using var doc = new XWPFDocument(stream);
        using var outStream = new MemoryStream();
        doc.write(outStream);
        outStream.Position = 0;

        using var archive = new ZipArchive(outStream, ZipArchiveMode.Read);
        var xml = ReadEntry(archive, "word/document.xml");
        AssertOrder(xml, "Cell before ", "<w:commentRangeStart", "commented cell ", "<w:commentRangeEnd", "<w:commentReference", "Cell after");
    }

    [Fact]
    public void RoundTrip_DocxCommentMarkers_PreservedNearHyperlinkAndTrackedChanges()
    {
        using var stream = CreateDocxWithDocumentXml("""
            <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                        xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
              <w:body>
                <w:p>
                  <w:r><w:t>Before </w:t></w:r>
                  <w:commentRangeStart w:id="8"/>
                  <w:hyperlink r:id="rLink"><w:r><w:t>link text</w:t></w:r></w:hyperlink>
                  <w:commentRangeEnd w:id="8"/>
                  <w:r><w:commentReference w:id="8"/></w:r>
                  <w:ins w:id="9" w:author="Reviewer"><w:r><w:t> inserted</w:t></w:r></w:ins>
                  <w:r><w:t> After</w:t></w:r>
                </w:p>
              </w:body>
            </w:document>
            """);

        using var doc = new XWPFDocument(stream);
        using var outStream = new MemoryStream();
        doc.write(outStream);
        outStream.Position = 0;

        using var archive = new ZipArchive(outStream, ZipArchiveMode.Read);
        var xml = ReadEntry(archive, "word/document.xml");
        AssertOrder(xml, "Before ", "<w:commentRangeStart", "link text", "<w:commentRangeEnd", "<w:commentReference", "<w:ins", " After");
        Assert.Contains("w:author=\"Reviewer\"", xml);
    }

    [Fact]
    public void Read_PoiDocxComments_ReturnsMinimalCommentModel()
    {
        using var input = new MemoryStream(File.ReadAllBytes(GetPoiTestDataPath("document/comment.docx")));
        using var doc = new XWPFDocument(input);

        var comment = Assert.Single(doc.getComments());
        Assert.Same(comment, doc.getCommentByID("0"));
        Assert.Equal("0", comment.getId());
        Assert.Equal("Unbekannter Autor", comment.getAuthor());
        Assert.Equal("", comment.getInitials());
        Assert.Equal("2019-10-11T05:43:39Z", comment.getDate());
        Assert.Equal("This is the first line\n\nThis is the second line", comment.getText());
        Assert.Null(doc.getCommentByID("missing"));
    }

    [Fact]
    public void Read_PoiDocxCommentMarkers_ReturnsParagraphAndRunReferenceIds()
    {
        using var input = new MemoryStream(File.ReadAllBytes(GetPoiTestDataPath("document/testComment.docx")));
        using var doc = new XWPFDocument(input);

        var paragraph = doc.getParagraphs().First(p => p.getText().Contains("comment ", StringComparison.Ordinal));

        Assert.Contains("0", paragraph.getCommentRangeStartIds());
        Assert.Contains("0", paragraph.getCommentRangeEndIds());
        Assert.Contains("0", paragraph.getCommentReferenceIds());
        Assert.Contains(paragraph.getRuns(), run => run.getCommentReferenceIds().Contains("0"));
    }

    [Fact]
    public void Write_DocxCommentApi_CreatesCommentsPartRelationshipAndRangeMarkers()
    {
        using var doc = new XWPFDocument();
        var paragraph = doc.createParagraph();
        paragraph.createRun().setText("Commented text");
        var comment = doc.createComment("Reviewer", "Created from dotnet-poi", "RV", "2026-05-08T08:30:00Z");
        paragraph.addComment(comment);

        using var stream = new MemoryStream();
        doc.write(stream);
        stream.Position = 0;

        using (var archive = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: true))
        {
            var documentXml = ReadEntry(archive, "word/document.xml");
            AssertOrder(documentXml, "<w:commentRangeStart w:id=\"0\"/>", "Commented text", "<w:commentRangeEnd w:id=\"0\"/>", "<w:commentReference w:id=\"0\"/>");

            var commentsXml = ReadEntry(archive, "word/comments.xml");
            Assert.Contains("<w:comment w:id=\"0\" w:author=\"Reviewer\" w:initials=\"RV\" w:date=\"2026-05-08T08:30:00Z\">", commentsXml);
            Assert.Contains("Created from dotnet-poi", commentsXml);

            var rels = ReadEntry(archive, "word/_rels/document.xml.rels");
            Assert.Contains("relationships/comments", rels);
            Assert.Contains("comments.xml", rels);

            var contentTypes = ReadEntry(archive, "[Content_Types].xml");
            Assert.Contains("/word/comments.xml", contentTypes);
            Assert.Contains("wordprocessingml.comments+xml", contentTypes);
        }

        stream.Position = 0;
        using var loaded = new XWPFDocument(stream);
        var loadedComment = Assert.Single(loaded.getComments());
        Assert.Equal("Reviewer", loadedComment.getAuthor());
        Assert.Equal("Created from dotnet-poi", loadedComment.getText());
        var loadedParagraph = Assert.Single(loaded.getParagraphs());
        Assert.Contains("0", loadedParagraph.getCommentRangeStartIds());
        Assert.Contains("0", loadedParagraph.getCommentRangeEndIds());
        Assert.Contains("0", loadedParagraph.getCommentReferenceIds());
    }

    [Fact]
    public void Write_DocxCommentApi_EditsExistingCommentsPart()
    {
        using var input = new MemoryStream(File.ReadAllBytes(GetPoiTestDataPath("document/comment.docx")));
        using var doc = new XWPFDocument(input);

        var comment = Assert.Single(doc.getComments());
        comment.setAuthor("Edited Author");
        comment.setText("Edited comment text");

        using var stream = new MemoryStream();
        doc.write(stream);
        stream.Position = 0;

        using var loaded = new XWPFDocument(stream);
        var loadedComment = Assert.Single(loaded.getComments());
        Assert.Equal("0", loadedComment.getId());
        Assert.Equal("Edited Author", loadedComment.getAuthor());
        Assert.Equal("Edited comment text", loadedComment.getText());
    }

    [Fact]
    public void Write_ParagraphsAndTables_KeepCreationOrder()
    {
        using var doc = new XWPFDocument();
        doc.createParagraph().createRun().setText("P1");
        var table = doc.createTable();
        table.createRow().createCell().addParagraph().createRun().setText("T1");
        doc.createParagraph().createRun().setText("P2");

        using var stream = new MemoryStream();
        doc.write(stream);
        stream.Position = 0;

        using var archive = new ZipArchive(stream, ZipArchiveMode.Read);
        var xml = ReadEntry(archive, "word/document.xml");

        var p1 = xml.IndexOf("P1", StringComparison.Ordinal);
        var t1 = xml.IndexOf("T1", StringComparison.Ordinal);
        var p2 = xml.IndexOf("P2", StringComparison.Ordinal);

        Assert.True(p1 >= 0);
        Assert.True(t1 > p1);
        Assert.True(p2 > t1);
    }

    [Fact]
    public void Read_InlineTextBoxContent_ExtractsTextOnParagraph()
    {
        using var stream = CreateDocxWithDocumentXml("""
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                        xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
                        xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                        xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
              <w:body>
                <w:p>
                  <w:r><w:t xml:space="preserve">Outer </w:t></w:r>
                  <w:r>
                    <w:drawing>
                      <wp:inline>
                        <a:graphic>
                          <a:graphicData>
                            <wps:wsp>
                              <wps:txbx>
                                <w:txbxContent>
                                  <w:p><w:r><w:t>Text box line 1</w:t></w:r></w:p>
                                  <w:p><w:r><w:t>Text box line 2</w:t></w:r></w:p>
                                </w:txbxContent>
                              </wps:txbx>
                            </wps:wsp>
                          </a:graphicData>
                        </a:graphic>
                      </wp:inline>
                    </w:drawing>
                  </w:r>
                </w:p>
                <w:sectPr/>
              </w:body>
            </w:document>
            """);

        using var doc = new XWPFDocument(stream);

        var para = Assert.Single(doc.getParagraphs());
        Assert.Equal("Outer Text box line 1\nText box line 2", para.getText());
    }

    [Fact]
    public void Read_AnchoredTextBoxContent_ExtractsTextWithoutCreatingBodyParagraphs()
    {
        using var stream = CreateDocxWithDocumentXml("""
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                        xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
                        xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                        xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
              <w:body>
                <w:p>
                  <w:r>
                    <w:drawing>
                      <wp:anchor>
                        <wp:docPr id="1" name="Text Box 1"/>
                        <a:graphic>
                          <a:graphicData>
                            <wps:wsp>
                              <wps:txbx>
                                <w:txbxContent>
                                  <w:p><w:r><w:t>Anchored box text</w:t></w:r></w:p>
                                </w:txbxContent>
                              </wps:txbx>
                            </wps:wsp>
                          </a:graphicData>
                        </a:graphic>
                      </wp:anchor>
                    </w:drawing>
                  </w:r>
                </w:p>
                <w:sectPr/>
              </w:body>
            </w:document>
            """);

        using var doc = new XWPFDocument(stream);

        var para = Assert.Single(doc.getParagraphs());
        Assert.Equal("Anchored box text", para.getText());
    }

    /// <summary>
    /// comments (word/comments.xml), footnotes (word/footnotes.xml), endnotes
    /// (word/endnotes.xml), OLE embeddings (word/embeddings/*) are separate
    /// ZIP parts NOT in GetModelEntryNames() → should be 🔵 preserved.
    /// (styles.xml preservation is already tested in RoundTrip_UnknownParts_Preserved.)
    /// </summary>
    [Fact]
    public void RoundTrip_CommentAndFootnoteSeparateParts_Preserved()
    {
        using var original = new XWPFDocument();
        original.createParagraph().createRun().setText("test");

        byte[] raw;
        using (var ms = new MemoryStream())
        {
            original.write(ms);
            raw = ms.ToArray();
        }

        using var injectedStream = new MemoryStream();
        using (var writerArchive = new ZipArchive(injectedStream, ZipArchiveMode.Create, leaveOpen: true))
        {
            using (var readerArchive = new ZipArchive(new MemoryStream(raw), ZipArchiveMode.Read))
            {
                foreach (var entry in readerArchive.Entries)
                {
                    using var s = entry.Open();
                    using var ms = new MemoryStream();
                    s.CopyTo(ms);
                    var ne = writerArchive.CreateEntry(entry.FullName);
                    using var ws = ne.Open();
                    ws.Write(ms.ToArray(), 0, (int)ms.Length);
                }
            }
            // Inject comments, footnotes, endnotes, and an OLE embedding
            var ce = writerArchive.CreateEntry("word/comments.xml");
            using (var ws = ce.Open())
                ws.Write(System.Text.Encoding.UTF8.GetBytes("<w:comments><w:comment /></w:comments>"));

            var fe = writerArchive.CreateEntry("word/footnotes.xml");
            using (var ws = fe.Open())
                ws.Write(System.Text.Encoding.UTF8.GetBytes("<w:footnotes><w:footnote /></w:footnotes>"));

            var ee = writerArchive.CreateEntry("word/endnotes.xml");
            using (var ws = ee.Open())
                ws.Write(System.Text.Encoding.UTF8.GetBytes("<w:endnotes><w:endnote /></w:endnotes>"));

            var oe = writerArchive.CreateEntry("word/embeddings/oleObject.bin");
            using (var ws = oe.Open())
                ws.Write(new byte[] { 0xDE, 0xAD, 0xBE, 0xEF });
        }

        injectedStream.Position = 0;
        using var loaded = new XWPFDocument(injectedStream);

        using var outStream = new MemoryStream();
        loaded.write(outStream);
        outStream.Position = 0;

        using var verify = new ZipArchive(outStream, ZipArchiveMode.Read);
        var names = verify.Entries.Select(e => e.FullName).ToHashSet(StringComparer.OrdinalIgnoreCase);
        Assert.Contains("word/comments.xml", names);
        Assert.Contains("word/footnotes.xml", names);
        Assert.Contains("word/endnotes.xml", names);
        Assert.Contains("word/embeddings/oleObject.bin", names);

        using var r = new StreamReader(verify.GetEntry("word/comments.xml")!.Open());
        Assert.Contains("<w:comments>", r.ReadToEnd());
    }

    [Fact]
    public void RoundTrip_TableCellMerge_Preserved()
    {
        using var original = new XWPFDocument();
        var table = original.createTable();
        table.addGridCol(9144);
        table.addGridCol(9144);
        table.addGridCol(9144);

        var row = table.createRow();
        var cell1 = row.createCell();
        cell1.addParagraph().createRun().setText("A1");
        cell1.setGridSpan(3); // merge 3 columns horizontally

        var row2 = table.createRow();
        var cell2 = row2.createCell();
        cell2.addParagraph().createRun().setText("B1");
        cell2.setVMerge("restart");
        var cell3 = row2.createCell();
        cell3.addParagraph().createRun().setText("B2");

        var row3 = table.createRow();
        var cell4 = row3.createCell();
        cell4.addParagraph().createRun().setText("C1");
        cell4.setVMerge("continue");
        var cell5 = row3.createCell();
        cell5.addParagraph().createRun().setText("C2");

        using var stream = new MemoryStream();
        original.write(stream);
        stream.Position = 0;

        using var loaded = new XWPFDocument(stream);
        var tables = loaded.getTables();
        Assert.Single(tables);
        var loadedTable = tables[0];
        Assert.Equal(3, loadedTable.getRows().Count);

        // Row 0: cells[0] has gridSpan=3
        var loadedRow0 = loadedTable.getRows()[0];
        var loadedCell = Assert.Single(loadedRow0.getCells());
        Assert.Equal(3, loadedCell.getGridSpan());
        Assert.Equal("A1", loadedCell.getParagraphs()[0].getText());

        // Row 1: cells[0] has vMerge="restart"
        var loadedRow1 = loadedTable.getRows()[1];
        Assert.Equal(2, loadedRow1.getCells().Count);
        Assert.Equal("restart", loadedRow1.getCells()[0].getVMerge());

        // Row 2: cells[0] has vMerge="continue"
        var loadedRow2 = loadedTable.getRows()[2];
        Assert.Equal(2, loadedRow2.getCells().Count);
        Assert.Equal("continue", loadedRow2.getCells()[0].getVMerge());
    }

    [Fact]
    public void RoundTrip_TableDepthApi_PreservesBordersCellWidthAndVerticalAlignment()
    {
        using var original = new XWPFDocument();
        var table = original.createTable();
        table.setWidth(9000, "dxa");
        table.setTopBorder(XWPFTable.XWPFBorderType.Single, 8, 0, "4472C4");
        table.setBottomBorder(XWPFTable.XWPFBorderType.Double, 12, 1, "70AD47");
        table.setInsideHBorder(XWPFTable.XWPFBorderType.Dotted, 4, 0, "auto");
        table.setInsideVBorder(XWPFTable.XWPFBorderType.Dashed, 6, 0, "FF0000");

        var row1 = table.createRow();
        row1.createCell().addParagraph().createRun().setText("A1");
        row1.createCell().addParagraph().createRun().setText("A2");
        row1.createCell().addParagraph().createRun().setText("A3");
        table.mergeCellsHorizontally(0, 0, 2);

        var row2 = table.createRow();
        var cell21 = row2.createCell();
        cell21.setWidth(3200, "dxa");
        cell21.setVerticalAlignment(XWPFTableCell.XWPFVertAlign.Center);
        cell21.addParagraph().createRun().setText("B1");
        row2.createCell().addParagraph().createRun().setText("B2");

        var row3 = table.createRow();
        var cell31 = row3.createCell();
        cell31.setWidth("33.3%");
        cell31.setVerticalAlignment(XWPFTableCell.XWPFVertAlign.Bottom);
        cell31.addParagraph().createRun().setText("C1");
        row3.createCell().addParagraph().createRun().setText("C2");
        table.mergeCellsVertically(1, 1, 2);

        using var stream = new MemoryStream();
        original.write(stream);
        stream.Position = 0;

        using (var archive = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: true))
        {
            var xml = ReadEntry(archive, "word/document.xml");
            Assert.Contains("<w:tblBorders>", xml);
            Assert.Contains("<w:top w:val=\"single\" w:sz=\"8\" w:space=\"0\" w:color=\"4472C4\"/>", xml);
            Assert.Contains("<w:bottom w:val=\"double\" w:sz=\"12\" w:space=\"1\" w:color=\"70AD47\"/>", xml);
            Assert.Contains("<w:insideH w:val=\"dotted\" w:sz=\"4\" w:space=\"0\" w:color=\"auto\"/>", xml);
            Assert.Contains("<w:gridSpan w:val=\"3\"/>", xml);
            Assert.Contains("<w:vAlign w:val=\"center\"/>", xml);
            Assert.Contains("<w:tcW w:w=\"1665\" w:type=\"pct\"/>", xml);
        }

        stream.Position = 0;
        using var loaded = new XWPFDocument(stream);
        var loadedTable = Assert.Single(loaded.getTables());
        Assert.Equal(XWPFTable.XWPFBorderType.Single, loadedTable.getTopBorderType());
        Assert.Equal(8, loadedTable.getTopBorderSize());
        Assert.Equal("4472C4", loadedTable.getTopBorderColor());
        Assert.Equal(XWPFTable.XWPFBorderType.Double, loadedTable.getBottomBorderType());
        Assert.Equal(1, loadedTable.getBottomBorderSpace());
        Assert.Equal(XWPFTable.XWPFBorderType.Dotted, loadedTable.getInsideHBorderType());
        Assert.Equal(XWPFTable.XWPFBorderType.Dashed, loadedTable.getInsideVBorderType());
        Assert.Equal("FF0000", loadedTable.getInsideVBorderColor());

        Assert.Equal(3, loadedTable.getRows()[0].getCells()[0].getGridSpan());
        Assert.Equal("continue", loadedTable.getRows()[0].getCells()[1].getHMerge());
        Assert.Equal("restart", loadedTable.getRows()[1].getCells()[1].getVMerge());
        Assert.Equal("continue", loadedTable.getRows()[2].getCells()[1].getVMerge());

        var loadedCell21 = loadedTable.getRows()[1].getCells()[0];
        Assert.Equal(3200, loadedCell21.getWidth());
        Assert.Equal("dxa", loadedCell21.getWidthType());
        Assert.Equal(XWPFTableCell.XWPFVertAlign.Center, loadedCell21.getVerticalAlignment());

        var loadedCell31 = loadedTable.getRows()[2].getCells()[0];
        Assert.Equal(1665, loadedCell31.getWidth());
        Assert.Equal("pct", loadedCell31.getWidthType());
        Assert.Equal(XWPFTableCell.XWPFVertAlign.Bottom, loadedCell31.getVerticalAlignment());
    }

    [Fact]
    public void Read_TableBordersFromExistingDocument_AvailableThroughApi()
    {
        using var stream = CreateDocxWithDocumentXml("""
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
              <w:body>
                <w:tbl>
                  <w:tblPr>
                    <w:tblW w:w="7200" w:type="dxa"/>
                    <w:tblBorders>
                      <w:top w:val="thick" w:sz="18" w:space="2" w:color="111111"/>
                      <w:left w:val="single" w:sz="8" w:space="0" w:color="222222"/>
                      <w:insideV w:val="dashed" w:sz="6" w:space="1" w:color="333333"/>
                    </w:tblBorders>
                  </w:tblPr>
                  <w:tr>
                    <w:tc>
                      <w:tcPr>
                        <w:tcW w:w="2400" w:type="dxa"/>
                        <w:vAlign w:val="bottom"/>
                      </w:tcPr>
                      <w:p><w:r><w:t>cell</w:t></w:r></w:p>
                    </w:tc>
                  </w:tr>
                </w:tbl>
                <w:sectPr/>
              </w:body>
            </w:document>
            """);

        using var doc = new XWPFDocument(stream);
        var table = Assert.Single(doc.getTables());
        Assert.Equal(7200, table.getWidth());
        Assert.Equal(XWPFTable.XWPFBorderType.Thick, table.getTopBorderType());
        Assert.Equal(18, table.getTopBorderSize());
        Assert.Equal(2, table.getTopBorderSpace());
        Assert.Equal("111111", table.getTopBorderColor());
        Assert.Equal(XWPFTable.XWPFBorderType.Single, table.getLeftBorderType());
        Assert.Equal(XWPFTable.XWPFBorderType.Dashed, table.getInsideVBorderType());
        Assert.Equal("333333", table.getInsideVBorderColor());

        var cell = table.getRows()[0].getCells()[0];
        Assert.Equal(2400, cell.getWidth());
        Assert.Equal("dxa", cell.getWidthType());
        Assert.Equal(XWPFTableCell.XWPFVertAlign.Bottom, cell.getVerticalAlignment());
    }

    [Fact]
    public void RoundTrip_ParagraphStyle_Preserved()
    {
        using var original = new XWPFDocument();
        var p = original.createParagraph();
        p.setStyle("Heading1");
        p.createRun().setText("Styled heading");

        using var stream = new MemoryStream();
        original.write(stream);
        stream.Position = 0;

        using var loaded = new XWPFDocument(stream);
        Assert.Single(loaded.getParagraphs());
        var loadedPara = loaded.getParagraphs()[0];
        Assert.Equal("Heading1", loadedPara.getStyleID());
        Assert.Equal("Styled heading", loadedPara.getText());
    }

    private static string ReadEntry(ZipArchive archive, string name)
    {
        var entry = archive.GetEntry(name);
        Assert.NotNull(entry);
        using var reader = new StreamReader(entry.Open());
        return reader.ReadToEnd();
    }

    private static void AssertOrder(string xml, params string[] fragments)
    {
        var previous = -1;
        foreach (var fragment in fragments)
        {
            var index = xml.IndexOf(fragment, StringComparison.Ordinal);
            Assert.True(index >= 0, $"Expected to find '{fragment}' in XML.");
            Assert.True(index > previous, $"Expected '{fragment}' to appear after the previous fragment.");
            previous = index;
        }
    }

    private static string GetPoiTestDataPath(string relativePath)
    {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (directory is not null)
        {
            var candidate = Path.Combine(directory.FullName, "poi/test-data", relativePath);
            if (File.Exists(candidate))
                return candidate;
            directory = directory.Parent;
        }

        throw new FileNotFoundException($"Could not locate POI test data file '{relativePath}'.");
    }

    private static MemoryStream CreateDocxWithDocumentXml(string documentXml)
    {
        var stream = new MemoryStream();
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true))
        {
            var entry = archive.CreateEntry("word/document.xml");
            using var writer = new StreamWriter(entry.Open());
            writer.Write(documentXml);
        }
        stream.Position = 0;
        return stream;
    }
}
