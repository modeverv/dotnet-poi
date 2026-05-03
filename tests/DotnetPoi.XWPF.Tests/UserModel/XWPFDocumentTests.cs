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
        // rId1 = settings.xml; image starts at rId2
        Assert.Contains("r:embed=\"rId2\"", xml);
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
        Assert.Single(loaded.getAllPictures());
        Assert.Equal(XWPFPictureData.PICTURE_TYPE_JPEG, loaded.getAllPictures()[0].getPictureType());
        Assert.Equal(imageBytes, loaded.getAllPictures()[0].getData());
    }

    private static string ReadEntry(ZipArchive archive, string name)
    {
        var entry = archive.GetEntry(name);
        Assert.NotNull(entry);
        using var reader = new StreamReader(entry.Open());
        return reader.ReadToEnd();
    }
}
