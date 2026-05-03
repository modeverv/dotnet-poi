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

    // ----- helpers -----

    private static XSLFPictureShape MakeShape()
    {
        using var prs = new XMLSlideShow();
        var slide = prs.createSlide();
        var idx   = prs.addPicture(new byte[] { 1, 2, 3 }, XSLFPictureData.PICTURE_TYPE_PNG);
        return prs.createPicture(slide, idx);
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
