using DotnetPoi.XSSF.UserModel;
using Xunit;

namespace DotnetPoi.Interop.Tests;

public class WriteForPoiTests
{
    private static readonly byte[] OneByOnePng =
        Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO2O8WcAAAAASUVORK5CYII=");

    [Fact]
    [Trait("Category", "WriteForPoi")]
    public void Write_StringAndNumberWorkbook_CreatesFixtureForPoi()
    {
        var fixturePath = GetFixturePath("phase0-basic.xlsx");
        Directory.CreateDirectory(Path.GetDirectoryName(fixturePath)!);

        using var workbook = new XSSFWorkbook();
        var sheet = workbook.createSheet("Phase0");
        var row = sheet.createRow(0);
        row.createCell(0).setCellValue("from dotnet-poi");
        row.createCell(1).setCellValue(123.25);
        row.createCell(2).setCellValue(0.0);

        using var stream = File.Create(fixturePath);
        workbook.write(stream);

        Assert.True(File.Exists(fixturePath));
        Assert.True(new FileInfo(fixturePath).Length > 0);
    }

    [Fact]
    [Trait("Category", "WriteForPoi")]
    public void Write_StyledWorkbook_CreatesFixtureForPoi()
    {
        var fixturePath = GetFixturePath("phase2-styles.xlsx");
        Directory.CreateDirectory(Path.GetDirectoryName(fixturePath)!);

        using var workbook = new XSSFWorkbook();
        var font = workbook.createFont();
        font.setBold(true);
        font.setItalic(true);
        font.setFontName("Arial");
        font.setFontHeightInPoints(14);
        font.setColor((short)IndexedColors.Red);

        var style = workbook.createCellStyle();
        style.setFont(font);
        style.setDataFormat(workbook.createDataFormat().getFormat("0.00"));
        style.setFillForegroundColor((short)IndexedColors.Yellow);
        style.setFillPattern(FillPatternType.SolidForeground);
        style.setBorderBottom(BorderStyle.Thin);

        var sheet = workbook.createSheet("Phase2");
        var cell = sheet.createRow(0).createCell(0);
        cell.setCellValue(123.456);
        cell.setCellStyle(style);

        using var stream = File.Create(fixturePath);
        workbook.write(stream);

        Assert.True(File.Exists(fixturePath));
        Assert.True(new FileInfo(fixturePath).Length > 0);
    }

    [Fact]
    [Trait("Category", "WriteForPoi")]
    public void Write_PictureWorkbook_CreatesFixtureForPoi()
    {
        var fixturePath = GetFixturePath("phase2_5-images.xlsx");
        Directory.CreateDirectory(Path.GetDirectoryName(fixturePath)!);

        using var workbook = new XSSFWorkbook();
        var sheet = workbook.createSheet("Phase2.5");
        sheet.createRow(0).createCell(0).setCellValue("image");

        var pictureIndex = workbook.addPicture(OneByOnePng, XSSFWorkbook.PICTURE_TYPE_PNG);
        var anchor = workbook.getCreationHelper().createClientAnchor();
        anchor.setCol1(0);
        anchor.setRow1(0);
        anchor.setCol2(1);
        anchor.setRow2(1);
        sheet.createDrawingPatriarch().createPicture(anchor, pictureIndex);

        using var stream = File.Create(fixturePath);
        workbook.write(stream);

        Assert.True(File.Exists(fixturePath));
        Assert.True(new FileInfo(fixturePath).Length > 0);
    }

    private static string GetFixturePath(string fileName)
    {
        var directory = AppContext.BaseDirectory;
        while (directory is not null)
        {
            var fixtures = Path.Combine(directory, "tests", "DotnetPoi.Interop.Tests", "fixtures");
            if (Directory.Exists(fixtures))
            {
                return Path.Combine(fixtures, "from-dotnet-poi", fileName);
            }

            directory = Directory.GetParent(directory)?.FullName;
        }

        throw new DirectoryNotFoundException("Could not locate tests/DotnetPoi.Interop.Tests/fixtures.");
    }
}
