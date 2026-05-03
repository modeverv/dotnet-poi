using System.IO.Compression;
using DotnetPoi.XSSF.UserModel;
using Xunit;

namespace DotnetPoi.XSSF.Tests.UserModel;

public class XSSFPictureTests
{
    private static readonly byte[] OneByOnePng =
        Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO2O8WcAAAAASUVORK5CYII=");

    [Fact]
    public void Write_PngPicture_ProducesDrawingPartsAndRelationships()
    {
        using var workbook = new XSSFWorkbook();
        var sheet = workbook.createSheet("Images");
        var pictureIndex = workbook.addPicture(OneByOnePng, XSSFWorkbook.PICTURE_TYPE_PNG);
        var drawing = sheet.createDrawingPatriarch();
        var anchor = workbook.getCreationHelper().createClientAnchor();
        anchor.setCol1(1);
        anchor.setRow1(2);
        anchor.setCol2(3);
        anchor.setRow2(4);

        drawing.createPicture(anchor, pictureIndex);

        using var stream = new MemoryStream();
        workbook.write(stream);

        stream.Position = 0;
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read);
        Assert.NotNull(archive.GetEntry("xl/media/image1.png"));
        Assert.NotNull(archive.GetEntry("xl/drawings/drawing1.xml"));
        Assert.NotNull(archive.GetEntry("xl/drawings/_rels/drawing1.xml.rels"));
        Assert.NotNull(archive.GetEntry("xl/worksheets/_rels/sheet1.xml.rels"));

        var contentTypes = ReadEntry(archive, "[Content_Types].xml");
        Assert.Contains("<Default Extension=\"png\" ContentType=\"image/png\"/>", contentTypes);
        Assert.Contains("<Override PartName=\"/xl/drawings/drawing1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.drawing+xml\"/>", contentTypes);

        var sheetXml = ReadEntry(archive, "xl/worksheets/sheet1.xml");
        Assert.Contains("<drawing r:id=\"rId1\"/>", sheetXml);

        var sheetRelationships = ReadEntry(archive, "xl/worksheets/_rels/sheet1.xml.rels");
        Assert.Contains("Target=\"../drawings/drawing1.xml\"", sheetRelationships);
        Assert.Contains("Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing\"", sheetRelationships);

        var drawingRelationships = ReadEntry(archive, "xl/drawings/_rels/drawing1.xml.rels");
        Assert.Contains("Target=\"../media/image1.png\"", drawingRelationships);
        Assert.Contains("Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\"", drawingRelationships);

        var drawingXml = ReadEntry(archive, "xl/drawings/drawing1.xml");
        Assert.Contains("<xdr:twoCellAnchor editAs=\"twoCell\">", drawingXml);
        Assert.Contains("<xdr:col>1</xdr:col>", drawingXml);
        Assert.Contains("<xdr:row>2</xdr:row>", drawingXml);
        Assert.Contains("<a:blip r:embed=\"rId1\"/>", drawingXml);
    }

    [Fact]
    public void Read_WorkbookWithMediaPart_RestoresPictureData()
    {
        using var original = new XSSFWorkbook();
        original.addPicture(OneByOnePng, XSSFWorkbook.PICTURE_TYPE_PNG);

        using var stream = new MemoryStream();
        original.write(stream);

        stream.Position = 0;
        using var loaded = new XSSFWorkbook(stream);

        var picture = Assert.Single(loaded.getAllPictures());
        Assert.Equal(XSSFWorkbook.PICTURE_TYPE_PNG, picture.getPictureType());
        Assert.Equal("png", picture.suggestFileExtension());
        Assert.Equal("image/png", picture.getMimeType());
        Assert.Equal(OneByOnePng, picture.getData());
    }

    private static string ReadEntry(ZipArchive archive, string name)
    {
        var entry = archive.GetEntry(name);
        Assert.NotNull(entry);
        using var reader = new StreamReader(entry.Open());
        return reader.ReadToEnd();
    }
}
