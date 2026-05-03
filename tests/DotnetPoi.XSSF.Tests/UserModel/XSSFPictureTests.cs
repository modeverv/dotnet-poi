using System.IO.Compression;
using DotnetPoi.XSSF.UserModel;
using Xunit;

namespace DotnetPoi.XSSF.Tests.UserModel;

public class XSSFPictureTests
{
    private static byte[] LoadTestImage() => File.ReadAllBytes("image.jpg");

    [Fact]
    public void Write_JpegPicture_ProducesDrawingPartsAndRelationships()
    {
        var imageBytes = LoadTestImage();

        using var workbook = new XSSFWorkbook();
        var sheet = workbook.createSheet("Images");
        var pictureIndex = workbook.addPicture(imageBytes, XSSFWorkbook.PICTURE_TYPE_JPEG);
        var drawing = sheet.createDrawingPatriarch();
        var anchor = workbook.getCreationHelper().createClientAnchor();
        anchor.setCol1(1);
        anchor.setRow1(2);
        anchor.setCol2(5);
        anchor.setRow2(15);

        drawing.createPicture(anchor, pictureIndex);

        using var stream = new MemoryStream();
        workbook.write(stream);

        stream.Position = 0;
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read);

        Assert.NotNull(archive.GetEntry("xl/media/image1.jpeg"));
        Assert.NotNull(archive.GetEntry("xl/drawings/drawing1.xml"));
        Assert.NotNull(archive.GetEntry("xl/drawings/_rels/drawing1.xml.rels"));
        Assert.NotNull(archive.GetEntry("xl/worksheets/_rels/sheet1.xml.rels"));

        var contentTypes = ReadEntry(archive, "[Content_Types].xml");
        Assert.Contains("<Default Extension=\"jpeg\" ContentType=\"image/jpeg\"/>", contentTypes);
        Assert.Contains("<Override PartName=\"/xl/drawings/drawing1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.drawing+xml\"/>", contentTypes);

        var sheetXml = ReadEntry(archive, "xl/worksheets/sheet1.xml");
        Assert.Contains("<drawing r:id=\"rId1\"/>", sheetXml);

        var sheetRelationships = ReadEntry(archive, "xl/worksheets/_rels/sheet1.xml.rels");
        Assert.Contains("Target=\"../drawings/drawing1.xml\"", sheetRelationships);
        Assert.Contains("Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing\"", sheetRelationships);

        var drawingRelationships = ReadEntry(archive, "xl/drawings/_rels/drawing1.xml.rels");
        Assert.Contains("Target=\"../media/image1.jpeg\"", drawingRelationships);
        Assert.Contains("Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\"", drawingRelationships);

        var drawingXml = ReadEntry(archive, "xl/drawings/drawing1.xml");
        Assert.Contains("<xdr:twoCellAnchor editAs=\"twoCell\">", drawingXml);
        Assert.Contains("<xdr:col>1</xdr:col>", drawingXml);
        Assert.Contains("<xdr:row>2</xdr:row>", drawingXml);
        Assert.Contains("<a:blip r:embed=\"rId1\"/>", drawingXml);

        // Verify media bytes are stored faithfully
        using var mediaEntry = archive.GetEntry("xl/media/image1.jpeg")!.Open();
        using var ms = new MemoryStream();
        mediaEntry.CopyTo(ms);
        Assert.Equal(imageBytes, ms.ToArray());
    }

    [Fact]
    public void Read_WorkbookWithJpegMedia_RestoresPictureData()
    {
        var imageBytes = LoadTestImage();

        using var original = new XSSFWorkbook();
        original.addPicture(imageBytes, XSSFWorkbook.PICTURE_TYPE_JPEG);

        using var stream = new MemoryStream();
        original.write(stream);

        stream.Position = 0;
        using var loaded = new XSSFWorkbook(stream);

        var picture = Assert.Single(loaded.getAllPictures());
        Assert.Equal(XSSFWorkbook.PICTURE_TYPE_JPEG, picture.getPictureType());
        Assert.Equal("jpeg", picture.suggestFileExtension());
        Assert.Equal("image/jpeg", picture.getMimeType());
        Assert.Equal(imageBytes, picture.getData());
    }

    [Fact]
    public void Write_MultipleSheets_EachWithOwnDrawing()
    {
        var imageBytes = LoadTestImage();

        using var workbook = new XSSFWorkbook();

        foreach (var sheetName in new[] { "Sheet1", "Sheet2" })
        {
            var sheet = workbook.createSheet(sheetName);
            var idx = workbook.addPicture(imageBytes, XSSFWorkbook.PICTURE_TYPE_JPEG);
            var anchor = workbook.getCreationHelper().createClientAnchor();
            anchor.setCol1(0); anchor.setRow1(0);
            anchor.setCol2(3); anchor.setRow2(5);
            sheet.createDrawingPatriarch().createPicture(anchor, idx);
        }

        using var stream = new MemoryStream();
        workbook.write(stream);

        stream.Position = 0;
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read);
        Assert.NotNull(archive.GetEntry("xl/drawings/drawing1.xml"));
        Assert.NotNull(archive.GetEntry("xl/drawings/drawing2.xml"));
        Assert.NotNull(archive.GetEntry("xl/media/image1.jpeg"));
        Assert.NotNull(archive.GetEntry("xl/media/image2.jpeg"));
    }

    private static string ReadEntry(ZipArchive archive, string name)
    {
        var entry = archive.GetEntry(name);
        Assert.NotNull(entry);
        using var reader = new StreamReader(entry.Open());
        return reader.ReadToEnd();
    }
}
