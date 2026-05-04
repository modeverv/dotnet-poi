using DotnetPoi.SS.UserModel;
using DotnetPoi.XSLF.UserModel;
using DotnetPoi.XSSF.UserModel;
using DotnetPoi.XWPF.UserModel;
using Xunit;

namespace DotnetPoi.Interop.Tests;

public class WriteForPoiTests
{
    private static byte[] LoadTestImage() => File.ReadAllBytes("image.jpg");

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

        var pictureIndex = workbook.addPicture(LoadTestImage(), XSSFWorkbook.PICTURE_TYPE_JPEG);
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

    [Fact]
    [Trait("Category", "WriteForPoi")]
    public void Write_RotatedPictureWorkbook_CreatesFixtureForPoi()
    {
        var fixturePath = GetFixturePath("phase3_1-rotation.xlsx");
        Directory.CreateDirectory(Path.GetDirectoryName(fixturePath)!);

        using var workbook = new XSSFWorkbook();
        var sheet = workbook.createSheet("Phase3.1");
        sheet.createRow(0).createCell(0).setCellValue("rotated");

        var pictureIndex = workbook.addPicture(LoadTestImage(), XSSFWorkbook.PICTURE_TYPE_JPEG);
        var anchor = workbook.getCreationHelper().createClientAnchor();
        anchor.setCol1(1);
        anchor.setRow1(1);
        anchor.setCol2(5);
        anchor.setRow2(15);
        var picture = sheet.createDrawingPatriarch().createPicture(anchor, pictureIndex);
        picture.setRotation(90.0);

        using var stream = File.Create(fixturePath);
        workbook.write(stream);

        Assert.True(File.Exists(fixturePath));
        Assert.True(new FileInfo(fixturePath).Length > 0);
    }

    [Fact]
    [Trait("Category", "WriteForPoi")]
    public void Write_DocxWithText_CreatesFixtureForPoi()
    {
        var fixturePath = GetFixturePath("phase3_2-basic.docx");
        Directory.CreateDirectory(Path.GetDirectoryName(fixturePath)!);

        using var doc = new XWPFDocument();
        var para1 = doc.createParagraph();
        var run1 = para1.createRun();
        run1.setText("from dotnet-poi docx");
        run1.setBold(true);

        var para2 = doc.createParagraph();
        var run2 = para2.createRun();
        run2.setText("second paragraph");
        run2.setItalic(true);

        using var stream = File.Create(fixturePath);
        doc.write(stream);

        Assert.True(File.Exists(fixturePath));
        Assert.True(new FileInfo(fixturePath).Length > 0);
    }

    [Fact]
    [Trait("Category", "WriteForPoi")]
    public void Write_DocxWithImageAndRotation_CreatesFixtureForPoi()
    {
        var fixturePath = GetFixturePath("phase3_2_1-image-rotation.docx");
        Directory.CreateDirectory(Path.GetDirectoryName(fixturePath)!);

        using var doc = new XWPFDocument();
        var para = doc.createParagraph();
        var run = para.createRun();
        var picture = run.addPicture(LoadTestImage(), XWPFPictureData.PICTURE_TYPE_JPEG, "image.jpg",
            914400, 914400);
        picture.setRotation(90.0);

        using var stream = File.Create(fixturePath);
        doc.write(stream);

        Assert.True(File.Exists(fixturePath));
        Assert.True(new FileInfo(fixturePath).Length > 0);
    }

    [Fact]
    [Trait("Category", "WriteForPoi")]
    public void Write_PptxWithPictureAndRotation_CreatesFixtureForPoi()
    {
        var fixturePath = GetFixturePath("phase3_3-pptx.pptx");
        Directory.CreateDirectory(Path.GetDirectoryName(fixturePath)!);

        using var prs = new XMLSlideShow();
        var slide = prs.createSlide();
        var picIdx = prs.addPicture(LoadTestImage(), XSLFPictureData.PICTURE_TYPE_JPEG);
        var shape = prs.createPicture(slide, picIdx);
        shape.setAnchor(0, 0, XMLSlideShow.DefaultSlideCx, XMLSlideShow.DefaultSlideCy);
        shape.setRotation(90.0);

        using var stream = File.Create(fixturePath);
        prs.write(stream);

        Assert.True(File.Exists(fixturePath));
        Assert.True(new FileInfo(fixturePath).Length > 0);
    }

    [Fact]
    [Trait("Category", "WriteForPoi")]
    public void Write_AgileEncryptedWorkbook_CreatesFixtureForPoi()
    {
        var fixturePath = GetFixturePath("phase3_4-agile-encrypted.xlsx");
        Directory.CreateDirectory(Path.GetDirectoryName(fixturePath)!);

        using var workbook = new XSSFWorkbook();
        var sheet = workbook.createSheet("Phase3.4");
        var row = sheet.createRow(0);
        row.createCell(0).setCellValue("encrypted from dotnet-poi");
        row.createCell(1).setCellValue(34.0);

        using var stream = File.Create(fixturePath);
        workbook.writeEncrypted(stream, "f");

        Assert.True(File.Exists(fixturePath));
        Assert.True(new FileInfo(fixturePath).Length > 0);
    }

    [Fact]
    [Trait("Category", "WriteForPoi")]
    public void Write_BooleanAndNumericCells_CreatesFixtureForPoi()
    {
        var fixturePath = GetFixturePath("phase7-cell-types.xlsx");
        Directory.CreateDirectory(Path.GetDirectoryName(fixturePath)!);

        using var workbook = new XSSFWorkbook();
        var sheet = workbook.createSheet("CellTypes");
        var row = sheet.createRow(0);
        row.createCell(0).setCellValue(true);           // BOOLEAN true
        row.createCell(1).setCellValue(false);          // BOOLEAN false
        row.createCell(2).setCellValue(42.5);           // NUMERIC
        row.createCell(3).setCellValue("hello");        // STRING

        using var stream = File.Create(fixturePath);
        workbook.write(stream);

        Assert.True(File.Exists(fixturePath));
        Assert.True(new FileInfo(fixturePath).Length > 0);
    }

    [Fact]
    [Trait("Category", "WriteForPoi")]
    public void Write_FormulaCells_CreatesFixtureForPoi()
    {
        var fixturePath = GetFixturePath("phase5-step1-formulas.xlsx");
        Directory.CreateDirectory(Path.GetDirectoryName(fixturePath)!);

        using var workbook = new XSSFWorkbook();
        var sheet = workbook.createSheet("Formulas");
        var row = sheet.createRow(0);
        row.createCell(0).setCellValue(10.0);
        row.createCell(1).setCellValue(20.0);

        var numericFormula = sheet.createRow(1).createCell(0);
        numericFormula.setCellFormula("A1+B1");

        var stringFormula = sheet.createRow(2).createCell(0);
        stringFormula.setCellFormula("\"hello \"&\"world\"");
        stringFormula.setCellValue("hello world");

        using var stream = File.Create(fixturePath);
        workbook.write(stream);

        Assert.True(File.Exists(fixturePath));
        Assert.True(new FileInfo(fixturePath).Length > 0);
    }

    [Fact]
    [Trait("Category", "WriteForPoi")]
    public void Write_ForceFormulaRecalculation_CreatesFixtureForPoi()
    {
        var fixturePath = GetFixturePath("phase5-step2-recalc.xlsx");
        Directory.CreateDirectory(Path.GetDirectoryName(fixturePath)!);

        using var workbook = new XSSFWorkbook();
        var sheet = workbook.createSheet("Recalc");
        sheet.createRow(0).createCell(0).setCellFormula("B1+C1");
        workbook.setForceFormulaRecalculation(true);

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
