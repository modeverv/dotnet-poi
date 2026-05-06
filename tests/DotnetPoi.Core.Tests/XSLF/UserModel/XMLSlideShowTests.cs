using System.IO.Compression;
using DotnetPoi.XSLF.UserModel;
using Xunit;

namespace DotnetPoi.XSLF.Tests.UserModel;

public class XMLSlideShowTests
{
    private static byte[] LoadTestImage() => File.ReadAllBytes("image.jpg");

    // ----- write tests -----

    [Fact]
    public void Write_EmptyPresentation_ProducesRequiredEntries()
    {
        using var prs = new XMLSlideShow();
        using var stream = new MemoryStream();
        prs.write(stream);

        stream.Position = 0;
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read);
        Assert.NotNull(archive.GetEntry("[Content_Types].xml"));
        Assert.NotNull(archive.GetEntry("_rels/.rels"));
        Assert.NotNull(archive.GetEntry("ppt/presentation.xml"));
        Assert.NotNull(archive.GetEntry("ppt/_rels/presentation.xml.rels"));
        Assert.NotNull(archive.GetEntry("ppt/slideMasters/slideMaster1.xml"));
        Assert.NotNull(archive.GetEntry("ppt/slideLayouts/slideLayout1.xml"));
    }

    [Fact]
    public void Write_SingleSlide_ProducesSlideEntry()
    {
        using var prs = new XMLSlideShow();
        prs.createSlide();
        using var stream = new MemoryStream();
        prs.write(stream);

        stream.Position = 0;
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read);
        Assert.NotNull(archive.GetEntry("ppt/slides/slide1.xml"));
        Assert.NotNull(archive.GetEntry("ppt/slides/_rels/slide1.xml.rels"));
    }

    [Fact]
    public void Write_MultipleSlides_ProducesOneEntryEach()
    {
        using var prs = new XMLSlideShow();
        prs.createSlide();
        prs.createSlide();
        prs.createSlide();
        using var stream = new MemoryStream();
        prs.write(stream);

        stream.Position = 0;
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read);
        Assert.NotNull(archive.GetEntry("ppt/slides/slide1.xml"));
        Assert.NotNull(archive.GetEntry("ppt/slides/slide2.xml"));
        Assert.NotNull(archive.GetEntry("ppt/slides/slide3.xml"));
    }

    [Fact]
    public void Write_SlideWithPicture_ProducesMediaAndRelationship()
    {
        var imageBytes = LoadTestImage();
        using var prs = new XMLSlideShow();
        var slide = prs.createSlide();
        var picIdx = prs.addPicture(imageBytes, XSLFPictureData.PICTURE_TYPE_JPEG);
        var shape = prs.createPicture(slide, picIdx);
        shape.setAnchor(0, 0, XMLSlideShow.DefaultSlideCx, XMLSlideShow.DefaultSlideCy);

        using var stream = new MemoryStream();
        prs.write(stream);

        stream.Position = 0;
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read);
        Assert.NotNull(archive.GetEntry("ppt/media/image1.jpeg"));

        var slideXml = ReadEntry(archive, "ppt/slides/slide1.xml");
        // rId1 = slideLayout; image starts at rId2
        Assert.Contains("r:embed=\"rId2\"", slideXml);
        Assert.Contains("<p:pic", slideXml);
    }

    [Fact]
    public void Write_PictureWithRotation_WritesRotAttribute()
    {
        var imageBytes = LoadTestImage();
        using var prs = new XMLSlideShow();
        var slide = prs.createSlide();
        var picIdx = prs.addPicture(imageBytes, XSLFPictureData.PICTURE_TYPE_JPEG);
        var shape = prs.createPicture(slide, picIdx);
        shape.setAnchor(0, 0, 914400, 914400);
        shape.setRotation(90.0);

        using var stream = new MemoryStream();
        prs.write(stream);

        stream.Position = 0;
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read);
        var slideXml = ReadEntry(archive, "ppt/slides/slide1.xml");
        // 90 degrees × 60000 = 5400000
        Assert.Contains("rot=\"5400000\"", slideXml);
    }

    [Fact]
    public void Write_PictureNoRotation_OmitsRotAttribute()
    {
        var imageBytes = LoadTestImage();
        using var prs = new XMLSlideShow();
        var slide = prs.createSlide();
        var picIdx = prs.addPicture(imageBytes, XSLFPictureData.PICTURE_TYPE_JPEG);
        var shape = prs.createPicture(slide, picIdx);
        shape.setAnchor(0, 0, 914400, 914400);

        using var stream = new MemoryStream();
        prs.write(stream);

        stream.Position = 0;
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read);
        var slideXml = ReadEntry(archive, "ppt/slides/slide1.xml");
        Assert.DoesNotContain("rot=", slideXml);
    }

    [Fact]
    public void Write_FlipHorizontal_WritesFlipHAttribute()
    {
        var imageBytes = LoadTestImage();
        using var prs = new XMLSlideShow();
        var slide = prs.createSlide();
        var picIdx = prs.addPicture(imageBytes, XSLFPictureData.PICTURE_TYPE_JPEG);
        var shape = prs.createPicture(slide, picIdx);
        shape.setAnchor(0, 0, 914400, 914400);
        shape.setFlipHorizontal(true);

        using var stream = new MemoryStream();
        prs.write(stream);

        stream.Position = 0;
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read);
        var slideXml = ReadEntry(archive, "ppt/slides/slide1.xml");
        Assert.Contains("flipH=\"1\"", slideXml);
    }

    [Fact]
    public void Write_SamePictureTwice_EmbedOnce()
    {
        var imageBytes = LoadTestImage();
        using var prs = new XMLSlideShow();
        var idx1 = prs.addPicture(imageBytes, XSLFPictureData.PICTURE_TYPE_JPEG);
        var idx2 = prs.addPicture(imageBytes, XSLFPictureData.PICTURE_TYPE_JPEG);
        Assert.Equal(idx1, idx2);
        Assert.Single(prs.getPictureData());
    }

    [Fact]
    public void AddPicture_ReturnsZeroBasedIndex()
    {
        var imageBytes = LoadTestImage();
        using var prs = new XMLSlideShow();
        var idx = prs.addPicture(imageBytes, XSLFPictureData.PICTURE_TYPE_JPEG);
        Assert.Equal(0, idx);
    }

    [Fact]
    public void PresentationXml_ContainsSlideSizeAndSlideId()
    {
        using var prs = new XMLSlideShow();
        prs.createSlide();
        using var stream = new MemoryStream();
        prs.write(stream);

        stream.Position = 0;
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read);
        var presXml = ReadEntry(archive, "ppt/presentation.xml");
        Assert.Contains("sldSz", presXml);
        Assert.Contains("sldId", presXml);
    }

    // ----- round-trip (write → read) tests -----

    [Fact]
    public void RoundTrip_SingleSlideNoPictures_SlideCountPreserved()
    {
        using var prs = new XMLSlideShow();
        prs.createSlide();

        using var stream = WriteToStream(prs);
        using var prs2 = new XMLSlideShow(stream);
        Assert.Single(prs2.getSlides());
    }

    [Fact]
    public void RoundTrip_MultipleSlides_CountPreserved()
    {
        using var prs = new XMLSlideShow();
        prs.createSlide();
        prs.createSlide();
        prs.createSlide();

        using var stream = WriteToStream(prs);
        using var prs2 = new XMLSlideShow(stream);
        Assert.Equal(3, prs2.getSlides().Count);
    }

    [Fact]
    public void RoundTrip_PictureEmbedded_ReadBackCorrectly()
    {
        var imageBytes = LoadTestImage();
        using var prs = new XMLSlideShow();
        var slide = prs.createSlide();
        var picIdx = prs.addPicture(imageBytes, XSLFPictureData.PICTURE_TYPE_JPEG);
        var shape = prs.createPicture(slide, picIdx);
        shape.setAnchor(100_000, 200_000, 914_400, 685_800);

        using var stream = WriteToStream(prs);
        using var prs2 = new XMLSlideShow(stream);

        Assert.Single(prs2.getSlides());
        var slide2 = prs2.getSlides()[0];
        Assert.Single(slide2.getShapes());
        var shape2 = slide2.getShapes()[0];
        Assert.Equal(XSLFPictureData.PICTURE_TYPE_JPEG, shape2.getPictureData().getPictureType());
        Assert.Equal(100_000L, shape2.getAnchorX());
        Assert.Equal(200_000L, shape2.getAnchorY());
        Assert.Equal(914_400L, shape2.getAnchorCx());
        Assert.Equal(685_800L, shape2.getAnchorCy());
    }

    [Fact]
    public void RoundTrip_RotationPreserved()
    {
        var imageBytes = LoadTestImage();
        using var prs = new XMLSlideShow();
        var slide = prs.createSlide();
        var picIdx = prs.addPicture(imageBytes, XSLFPictureData.PICTURE_TYPE_JPEG);
        var shape = prs.createPicture(slide, picIdx);
        shape.setAnchor(0, 0, 914400, 914400);
        shape.setRotation(45.0);

        using var stream = WriteToStream(prs);
        using var prs2 = new XMLSlideShow(stream);
        var shape2 = prs2.getSlides()[0].getShapes()[0];
        Assert.Equal(45.0, shape2.getRotation(), precision: 3);
    }

    [Fact]
    public void RoundTrip_PictureBytes_Preserved()
    {
        var imageBytes = LoadTestImage();
        using var prs = new XMLSlideShow();
        var slide = prs.createSlide();
        var picIdx = prs.addPicture(imageBytes, XSLFPictureData.PICTURE_TYPE_JPEG);
        prs.createPicture(slide, picIdx).setAnchor(0, 0, 914400, 914400);

        using var stream = WriteToStream(prs);
        using var prs2 = new XMLSlideShow(stream);
        var picData = prs2.getSlides()[0].getShapes()[0].getPictureData();
        Assert.Equal(imageBytes, picData.getData());
    }

    // ----- XSLFPictureShape unit tests -----

    [Fact]
    public void SetRotation_360_NormalisesToZero()
    {
        var shape = MakeShape();
        shape.setRotation(360.0);
        Assert.Equal(0.0, shape.getRotation(), precision: 3);
    }

    [Fact]
    public void SetRotation_Negative_NormalisesPositive()
    {
        var shape = MakeShape();
        shape.setRotation(-90.0);
        Assert.Equal(270.0, shape.getRotation(), precision: 3);
    }

    [Fact]
    public void SetRotation_RoundTrip_180Degrees()
    {
        var shape = MakeShape();
        shape.setRotation(180.0);
        Assert.Equal(180.0, shape.getRotation(), precision: 3);
        Assert.Equal(10_800_000, shape.getRotationAttribute());
    }

    // ----- text box round-trip tests -----

    [Fact]
    public void RoundTrip_TextPreserved_SingleRun()
    {
        using var prs = new XMLSlideShow();
        var slide = prs.createSlide();
        var textBox = slide.createTextBox();
        textBox.setAnchor(100_000, 200_000, 914_400, 685_800);
        var para = textBox.addParagraph();
        para.addRun("Hello PPTX");

        using var stream = WriteToStream(prs);
        using var prs2 = new XMLSlideShow(stream);

        var slide2 = prs2.getSlides()[0];
        Assert.Single(slide2.getAutoShapes());
        var tb2 = slide2.getAutoShapes()[0];
        Assert.Single(tb2.Paragraphs);
        Assert.Single(tb2.Paragraphs[0].Runs);
        Assert.Equal("Hello PPTX", tb2.Paragraphs[0].Runs[0].Text);
    }

    [Fact]
    public void RoundTrip_TextPreserved_MultiRun()
    {
        using var prs = new XMLSlideShow();
        var slide = prs.createSlide();
        var textBox = slide.createTextBox();
        textBox.setAnchor(0, 0, 914_400, 685_800);
        var para = textBox.addParagraph();
        para.addRun("Bold ").Bold = true;
        para.addRun("Normal");

        using var stream = WriteToStream(prs);
        using var prs2 = new XMLSlideShow(stream);

        var slide2 = prs2.getSlides()[0];
        var tb2 = slide2.getAutoShapes()[0];
        Assert.Equal(2, tb2.Paragraphs[0].Runs.Count);
        Assert.True(tb2.Paragraphs[0].Runs[0].Bold);
        Assert.Equal("Bold ", tb2.Paragraphs[0].Runs[0].Text);
        Assert.Equal("Normal", tb2.Paragraphs[0].Runs[1].Text);
    }

    [Fact]
    public void RoundTrip_MultipleParagraphs_Preserved()
    {
        using var prs = new XMLSlideShow();
        var slide = prs.createSlide();
        var textBox = slide.createTextBox();
        textBox.setAnchor(0, 0, 914_400, 685_800);
        textBox.addParagraph().addRun("First paragraph");
        textBox.addParagraph().addRun("Second paragraph");

        using var stream = WriteToStream(prs);
        using var prs2 = new XMLSlideShow(stream);

        var tb2 = prs2.getSlides()[0].getAutoShapes()[0];
        Assert.Equal(2, tb2.Paragraphs.Count);
        Assert.Equal("First paragraph", tb2.Paragraphs[0].Runs[0].Text);
        Assert.Equal("Second paragraph", tb2.Paragraphs[1].Runs[0].Text);
    }

    [Fact]
    public void RoundTrip_AutoShapeAnchor_Preserved()
    {
        using var prs = new XMLSlideShow();
        var slide = prs.createSlide();
        var textBox = slide.createTextBox();
        textBox.setAnchor(500_000, 1_000_000, 2_000_000, 3_000_000);
        textBox.addParagraph().addRun("test");

        using var stream = WriteToStream(prs);
        using var prs2 = new XMLSlideShow(stream);

        var tb2 = prs2.getSlides()[0].getAutoShapes()[0];
        Assert.Equal(500_000L, tb2.getAnchorX());
        Assert.Equal(1_000_000L, tb2.getAnchorY());
        Assert.Equal(2_000_000L, tb2.getAnchorCx());
        Assert.Equal(3_000_000L, tb2.getAnchorCy());
    }

    [Fact]
    public void RoundTrip_SlideSize_Preserved()
    {
        long customCx = 6_858_000L;
        long customCy = 5_143_500L;
        using var prs = new XMLSlideShow();
        prs.setSlideSize(customCx, customCy);
        prs.createSlide();

        using var stream = WriteToStream(prs);
        using var prs2 = new XMLSlideShow(stream);

        Assert.Equal(customCx, prs2.getSlideCx());
        Assert.Equal(customCy, prs2.getSlideCy());
    }

    [Fact]
    public void RoundTrip_TextBoxAndPicture_Combined()
    {
        var image = LoadTestImage();
        using var prs = new XMLSlideShow();
        var slide = prs.createSlide();
        var picIdx = prs.addPicture(image, XSLFPictureData.PICTURE_TYPE_JPEG);
        prs.createPicture(slide, picIdx).setAnchor(0, 0, 914_400, 685_800);
        var tb = slide.createTextBox();
        tb.setAnchor(100_000, 100_000, 500_000, 300_000);
        tb.addParagraph().addRun("Hello");

        using var stream = WriteToStream(prs);
        using var prs2 = new XMLSlideShow(stream);

        Assert.Single(prs2.getSlides()[0].getShapes());
        Assert.Single(prs2.getSlides()[0].getAutoShapes());
        Assert.Equal("Hello", prs2.getSlides()[0].getAutoShapes()[0].Paragraphs[0].Runs[0].Text);
    }

    // ----- helpers -----

    private static XSLFPictureShape MakeShape()
    {
        using var prs = new XMLSlideShow();
        var slide = prs.createSlide();
        var idx   = prs.addPicture(new byte[] { 1, 2, 3 }, XSLFPictureData.PICTURE_TYPE_PNG);
        return prs.createPicture(slide, idx);
    }

    [Fact]
    public void RoundTrip_UnknownParts_Preserved()
    {
        // Prepare a pptx bytes with extra entries not in the model's known list
        using var original = new XMLSlideShow();
        var slide = original.createSlide();
        slide.createTextBox();
        slide.getAutoShapes()[0].setAnchor(0, 0, 100000, 100000);
        slide.getAutoShapes()[0].addParagraph().addRun("test");

        byte[] raw;
        using (var ms = new MemoryStream())
        {
            original.write(ms);
            raw = ms.ToArray();
        }

        // Re-pack the zip with additional non-model entries
        using var injectedStream = new MemoryStream();
        using (var writerArchive = new ZipArchive(injectedStream, ZipArchiveMode.Create, leaveOpen: true))
        {
            // Copy all existing entries
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
            // Add extra non-model entries
            var extraEntry = writerArchive.CreateEntry("ppt/slideLayouts/layout2.xml");
            using (var ws = extraEntry.Open())
            {
                ws.Write(System.Text.Encoding.UTF8.GetBytes("<p:sldLayout/>"));
            }
            extraEntry = writerArchive.CreateEntry("ppt/slideLayouts/_rels/layout2.xml.rels");
            using (var ws = extraEntry.Open())
            {
                ws.Write(System.Text.Encoding.UTF8.GetBytes("<?xml version=\"1.0\"?><Relationships/>"));
            }
        }
        injectedStream.Position = 0;

        // Load the injected pptx — CollectPreservedEntries should capture the extra layout
        using var loaded = new XMLSlideShow(injectedStream);
        Assert.Single(loaded.getSlides());

        // Write again — preserved entries should be emitted
        using var reloadedStream = new MemoryStream();
        loaded.write(reloadedStream);
        reloadedStream.Position = 0;

        // Verify extra entries survive the round-trip
        using var finalArchive = new ZipArchive(reloadedStream, ZipArchiveMode.Read);
        var layoutEntry = finalArchive.GetEntry("ppt/slideLayouts/layout2.xml");
        Assert.NotNull(layoutEntry);
        using var layoutReader = new StreamReader(layoutEntry.Open());
        Assert.Equal("<p:sldLayout/>", layoutReader.ReadToEnd());

        var relsEntry = finalArchive.GetEntry("ppt/slideLayouts/_rels/layout2.xml.rels");
        Assert.NotNull(relsEntry);
    }

    [Fact]
    public void RoundTrip_Table_Restored()
    {
        using var original = new XMLSlideShow();
        var slide = original.createSlide();

        // Create a table with 2 columns, 2 rows
        var table = slide.createTable();
        table.setAnchor(100000, 200000, 3000000, 1000000);
        table.addGridCol(1500000);
        table.addGridCol(1500000);

        var row1 = table.createRow();
        var cellA1 = row1.createCell();
        cellA1.addParagraph().addRun("A1");
        var cellB1 = row1.createCell();
        cellB1.addParagraph().addRun("B1");

        var row2 = table.createRow();
        var cellA2 = row2.createCell();
        cellA2.addParagraph().addRun("A2");
        var cellB2 = row2.createCell();
        cellB2.addParagraph().addRun("B2");

        using var stream = WriteToStream(original);
        using var loaded = new XMLSlideShow(stream);

        var loadedSlide = loaded.getSlides()[0];
        var tables = loadedSlide.getTables();
        Assert.Single(tables);

        var loadedTable = tables[0];
        Assert.Equal(2, loadedTable.GridColWidths.Count);
        Assert.Equal(1500000, loadedTable.GridColWidths[0]);
        Assert.Equal(1500000, loadedTable.GridColWidths[1]);
        Assert.Equal(2, loadedTable.Rows.Count);
        Assert.Equal(2, loadedTable.Rows[0].Cells.Count);
        Assert.Equal(2, loadedTable.Rows[1].Cells.Count);

        Assert.Equal("A1", loadedTable.Rows[0].Cells[0].Paragraphs[0].getPlainText());
        Assert.Equal("B1", loadedTable.Rows[0].Cells[1].Paragraphs[0].getPlainText());
        Assert.Equal("A2", loadedTable.Rows[1].Cells[0].Paragraphs[0].getPlainText());
        Assert.Equal("B2", loadedTable.Rows[1].Cells[1].Paragraphs[0].getPlainText());

        // Verify anchor preserved
        Assert.Equal(100000, loadedTable.getAnchorX());
        Assert.Equal(200000, loadedTable.getAnchorY());
        Assert.Equal(3000000, loadedTable.getAnchorCx());
        Assert.Equal(1000000, loadedTable.getAnchorCy());
    }

    /// <summary>
    /// Notes slides (ppt/notesSlides/notesSlide1.xml) and non-image media
    /// (video/audio in ppt/media/ not tracked in _pictures) are separate ZIP
    /// parts NOT in GetModelEntryNames() → should be 🔵 preserved.
    /// </summary>
    [Fact]
    public void RoundTrip_NotesSlides_Preserved()
    {
        using var original = new XMLSlideShow();
        original.createSlide();

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
            // Inject notes slide
            var notesEntry = writerArchive.CreateEntry("ppt/notesSlides/notesSlide1.xml");
            using (var ws = notesEntry.Open())
                ws.Write(System.Text.Encoding.UTF8.GetBytes("<p:notes/>"));
            // Inject a non-image media file (simulates video/audio)
            var mediaEntry = writerArchive.CreateEntry("ppt/media/video1.mp4");
            using (var ws = mediaEntry.Open())
                ws.Write(new byte[] { 0x00, 0x01, 0x02 });
        }

        injectedStream.Position = 0;
        using var loaded = new XMLSlideShow(injectedStream);
        using var outStream = new MemoryStream();
        loaded.write(outStream);
        outStream.Position = 0;

        using var verify = new ZipArchive(outStream, ZipArchiveMode.Read);
        var names = verify.Entries.Select(e => e.FullName).ToHashSet(StringComparer.OrdinalIgnoreCase);
        Assert.Contains("ppt/notesSlides/notesSlide1.xml", names);
        Assert.Contains("ppt/media/video1.mp4", names);

        using var r = new StreamReader(verify.GetEntry("ppt/notesSlides/notesSlide1.xml")!.Open());
        var notesContent = r.ReadToEnd();
        Assert.Contains("p:notes", notesContent, StringComparison.OrdinalIgnoreCase);
    }

    private static MemoryStream WriteToStream(XMLSlideShow prs)
    {
        var stream = new MemoryStream();
        prs.write(stream);
        stream.Position = 0;
        return stream;
    }

    private static string ReadEntry(ZipArchive archive, string name)
    {
        var entry = archive.GetEntry(name);
        Assert.NotNull(entry);
        using var s = entry.Open();
        using var r = new StreamReader(s);
        return r.ReadToEnd();
    }
}
