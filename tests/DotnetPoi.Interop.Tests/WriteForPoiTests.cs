using DotnetPoi.SS.UserModel;
using DotnetPoi.SS.Util;
using DotnetPoi.HSSF.UserModel;
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
    public void Write_HssfWorkbook_CreatesFixtureForPoi()
    {
        var fixturePath = GetFixturePath("phase6-basic.xls");
        Directory.CreateDirectory(Path.GetDirectoryName(fixturePath)!);

        using var workbook = new HSSFWorkbook();
        var sheet = workbook.createSheet("Phase6");
        var row = sheet.createRow(0);
        row.createCell(0).setCellValue("from dotnet-poi hssf");
        row.createCell(1).setCellValue(66.25);
        row.createCell(2).setCellValue(true);

        using var stream = File.Create(fixturePath);
        workbook.write(stream);
        stream.Flush();

        Assert.True(File.Exists(fixturePath));
        Assert.True(new FileInfo(fixturePath).Length > 0);
    }

    [Fact]
    [Trait("Category", "WriteForPoi")]
    public void Write_Phase12HssfStyles_CreatesFixtureForPoi()
    {
        // Phase 12 item 4: Direction B — dotnet-poi writes .xls with styles → Java POI reads
        var fixturePath = GetFixturePath("phase12-hssf-styles.xls");
        Directory.CreateDirectory(Path.GetDirectoryName(fixturePath)!);

        using var workbook = new HSSFWorkbook();

        var boldFont = workbook.createFont();
        boldFont.setBold(true);
        boldFont.setFontName("Calibri");
        boldFont.setFontHeightInPoints(14);
        boldFont.setItalic(true);

        var style1 = workbook.createCellStyle();
        style1.setFont(boldFont);
        style1.setAlignment(HorizontalAlignment.Center);
        style1.setWrapText(true);
        style1.setBorderBottom(BorderStyle.Thin);
        style1.setDataFormat(workbook.createDataFormat().getFormat("0.00"));

        var style2 = workbook.createCellStyle();
        style2.setAlignment(HorizontalAlignment.Right);
        style2.setBorderLeft(BorderStyle.Medium);

        var sheet = workbook.createSheet("Styles");
        var row = sheet.createRow(0);
        var cell0 = row.createCell(0);
        cell0.setCellValue(42.5);
        cell0.setCellStyle(style1);

        var cell1 = row.createCell(1);
        cell1.setCellValue("right");
        cell1.setCellStyle(style2);

        row.createCell(2).setCellValue("no style");

        using var stream = File.Create(fixturePath);
        workbook.write(stream);
        stream.Flush();

        Assert.True(File.Exists(fixturePath));
        Assert.True(new FileInfo(fixturePath).Length > 0);
    }

    [Fact]
    [Trait("Category", "WriteForPoi")]
    public void Write_Phase12HssfLayout_CreatesFixtureForPoi()
    {
        // Phase 12 item 4: Direction B — dotnet-poi writes .xls with layout → Java POI reads
        var fixturePath = GetFixturePath("phase12-hssf-layout.xls");
        Directory.CreateDirectory(Path.GetDirectoryName(fixturePath)!);

        using var workbook = new HSSFWorkbook();
        var sheet = workbook.createSheet("Layout");

        sheet.setColumnWidth(0, 5000);
        sheet.setColumnWidth(1, 8000);
        sheet.setColumnHidden(2, true);

        var row0 = sheet.createRow(0);
        row0.setHeight(30.0f);
        row0.createCell(0).setCellValue("wide col");

        var row1 = sheet.createRow(1);
        row1.setHidden(true);
        row1.createCell(0).setCellValue("hidden row");

        var row2 = sheet.createRow(2);
        row2.setHeight(20.0f);
        row2.createCell(0).setCellValue("normal");

        sheet.addMergedRegion(new SS.Util.CellRangeAddress(3, 3, 0, 2));
        sheet.createRow(3).createCell(0).setCellValue("merged");

        sheet.createFreezePane(1, 2);

        using var stream = File.Create(fixturePath);
        workbook.write(stream);
        stream.Flush();

        Assert.True(File.Exists(fixturePath));
        Assert.True(new FileInfo(fixturePath).Length > 0);
    }

    [Fact]
    [Trait("Category", "WriteForPoi")]
    public void Write_Phase12HssfUnicode_CreatesFixtureForPoi()
    {
        // Phase 12 item 3: Direction B — Unicode/Japanese sheet names and strings
        var fixturePath = GetFixturePath("phase12-hssf-unicode.xls");
        Directory.CreateDirectory(Path.GetDirectoryName(fixturePath)!);

        using var workbook = new HSSFWorkbook();

        var sheet1 = workbook.createSheet("日本語");
        var row0 = sheet1.createRow(0);
        row0.createCell(0).setCellValue("テスト文字列");
        row0.createCell(1).setCellValue("hello 世界");
        row0.createCell(2).setCellValue("こんにちは");

        var sheet2 = workbook.createSheet("中文测试");
        sheet2.createRow(0).createCell(0).setCellValue("汉字测试");

        using var stream = File.Create(fixturePath);
        workbook.write(stream);
        stream.Flush();

        Assert.True(File.Exists(fixturePath));
        Assert.True(new FileInfo(fixturePath).Length > 0);
    }

    [Fact]
    [Trait("Category", "WriteForPoi")]
    public void Write_Phase12HssfComprehensive_CreatesFixtureForPoi()
    {
        // Phase 12 item 3: Direction B — dotnet-poi writes .xls → Java POI reads
        var fixturePath = GetFixturePath("phase12-hssf-comprehensive.xls");
        Directory.CreateDirectory(Path.GetDirectoryName(fixturePath)!);

        using var workbook = new HSSFWorkbook();

        // Sheet 1: all cell types
        var sheet1 = workbook.createSheet("CellTypes");
        var row0 = sheet1.createRow(0);
        row0.createCell(0).setCellValue("string value");
        row0.createCell(1).setCellValue(42.5);
        row0.createCell(2).setCellValue(true);
        row0.createCell(3).setCellValue(false);
        row0.createCell(4).setCellErrorValue(0x07);  // #DIV/0!
        row0.createCell(5).setCellErrorValue(0x2A);  // #N/A
        row0.createCell(6);                          // blank

        // Sheet 2: sparse layout
        var sheet2 = workbook.createSheet("Sparse");
        sheet2.createRow(0).createCell(0).setCellValue("row0col0");
        sheet2.createRow(5).createCell(3).setCellValue(99.9);
        sheet2.createRow(10).createCell(0).setCellValue("row10");

        using var stream = File.Create(fixturePath);
        workbook.write(stream);
        stream.Flush();

        Assert.True(File.Exists(fixturePath));
        Assert.True(new FileInfo(fixturePath).Length > 0);
    }

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
    public void Write_ComprehensiveDocx_CreatesFixtureForPoi()
    {
        var fixturePath = GetFixturePath("phase-docx-comprehensive.docx");
        Directory.CreateDirectory(Path.GetDirectoryName(fixturePath)!);

        using var doc = new XWPFDocument();

        // --- Page setup ---
        doc.setPageSize(15840, 12240); // landscape A4-ish
        doc.setLandscape(true);
        doc.setMargins(720, 720, 720, 720);

        // --- Header / Footer ---
        doc.setHeaderText("Interop Test Header");
        doc.setFooterText("Page");

        // --- 1) Rich text paragraph (multiple runs with formatting) ---
        var richPara = doc.createParagraph();
        var runBold = richPara.createRun();
        runBold.setText("Bold ");
        runBold.setBold(true);
        runBold.setFontSize(14);
        runBold.setFontName("Arial");
        runBold.setColor("FF0000");

        var runItalic = richPara.createRun();
        runItalic.setText("Italic ");
        runItalic.setItalic(true);
        runItalic.setFontSize(12);
        runItalic.setColor("0000FF");

        var runUnderline = richPara.createRun();
        runUnderline.setText("Underline ");
        runUnderline.setUnderline(true);
        runUnderline.setFontName("Times New Roman");

        var runStrike = richPara.createRun();
        runStrike.setText("Strikethrough");
        runStrike.setStrike(true);
        runStrike.setFontSize(16);

        // --- 2) Numbered list ---
        var numPara1 = doc.createParagraph();
        numPara1.setNumberedList();
        numPara1.createRun().setText("First item");

        var numPara2 = doc.createParagraph();
        numPara2.setNumberedList();
        numPara2.createRun().setText("Second item");

        // --- 3) Bullet list ---
        var bullet1 = doc.createParagraph();
        bullet1.setBulletList();
        bullet1.createRun().setText("Bullet A");

        var bullet2 = doc.createParagraph();
        bullet2.setBulletList();
        bullet2.createRun().setText("Bullet B");

        // --- 4) Indentation and spacing ---
        var indentPara = doc.createParagraph();
        indentPara.setAlignment(ParagraphAlignment.Center);
        indentPara.setIndentationLeft(720);   // 0.5in
        indentPara.setIndentationRight(360);  // 0.25in
        indentPara.setIndentationFirstLine(720);
        indentPara.setSpacingBefore(240);
        indentPara.setSpacingAfter(120);
        indentPara.createRun().setText("Indented centered paragraph with spacing before and after.");

        // --- 5) Hyperlink (external URL) ---
        var linkPara = doc.createParagraph();
        var linkRun = linkPara.createRun();
        linkRun.setText("Click here for DotnetPoi");
        linkRun.setHyperlink("https://github.com/dotnetpoi/DotnetPoi");

        // --- 6) Table with row / cell text ---
        var table = doc.createTable();
        table.addGridCol(4572);
        table.addGridCol(4572);
        table.addGridCol(4572);

        var headerRow = table.createRow();
        headerRow.createCell().addParagraph().createRun().setText("Col A");
        headerRow.createCell().addParagraph().createRun().setText("Col B");
        headerRow.createCell().addParagraph().createRun().setText("Col C");

        var dataRow = table.createRow();
        dataRow.createCell().addParagraph().createRun().setText("A1");
        dataRow.createCell().addParagraph().createRun().setText("B1");
        dataRow.createCell().addParagraph().createRun().setText("C1");

        // --- 7) Hyperlink in a table cell ---
        var linkRow = table.createRow();
        linkRow.createCell().addParagraph().createRun().setText("Link cell");
        var linkCell = linkRow.createCell().addParagraph().createRun();
        linkCell.setText("Example");
        linkCell.setHyperlink("https://example.com");
        linkRow.createCell().addParagraph().createRun().setText("End");

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
    public void Write_PptxWithTextBoxesAndTables_CreatesFixtureForPoi()
    {
        var fixturePath = GetFixturePath("phase-pptx-comprehensive.pptx");
        Directory.CreateDirectory(Path.GetDirectoryName(fixturePath)!);

        using var prs = new XMLSlideShow();
        var slide = prs.createSlide();

        // --- Text box with formatted text ---
        var tb = slide.createTextBox();
        tb.setAnchor(100000, 100000, 4000000, 500000);
        var para = tb.addParagraph();
        var run = para.addRun("Bold text");
        run.Bold = true;
        run.FontSize = 18;
        var para2 = tb.addParagraph();
        var run2 = para2.addRun("Italic text");
        run2.Italic = true;
        run2.FontSize = 14;

        // --- Table with rows and cells ---
        var table = slide.createTable();
        table.setAnchor(100000, 700000, 3000000, 1500000);
        table.addGridCol(1500000);
        table.addGridCol(1500000);
        var row1 = table.createRow();
        var cellA1 = row1.createCell();
        cellA1.addParagraph().addRun("Cell A1");
        var cellB1 = row1.createCell();
        cellB1.addParagraph().addRun("Cell B1");
        var row2 = table.createRow();
        var cellA2 = row2.createCell();
        cellA2.addParagraph().addRun("Cell A2");
        var cellB2 = row2.createCell();
        cellB2.addParagraph().addRun("Cell B2");

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

    [Fact]
    [Trait("Category", "WriteForPoi")]
    public void Write_EvaluatedFormulaFunctions_CreatesFixtureForPoi()
    {
        var fixturePath = GetFixturePath("phase5-step3-evaluated-functions.xlsx");
        Directory.CreateDirectory(Path.GetDirectoryName(fixturePath)!);

        using var workbook = new XSSFWorkbook();
        var sheet = workbook.createSheet("Eval");
        var row = sheet.createRow(0);
        row.createCell(0).setCellValue(10.0);
        row.createCell(1).setCellValue(20.0);
        row.createCell(2).setCellValue(30.0);

        var sum = sheet.createRow(1).createCell(0);
        sum.setCellFormula("SUM(A1:C1)");
        var average = sheet.getRow(1)!.createCell(1);
        average.setCellFormula("AVERAGE(A1:C1)");
        var text = sheet.getRow(1)!.createCell(2);
        text.setCellFormula("CONCATENATE(\"sum=\",SUM(A1:C1))");

        workbook.getCreationHelper().createFormulaEvaluator().evaluateAll();

        using var stream = File.Create(fixturePath);
        workbook.write(stream);

        Assert.True(File.Exists(fixturePath));
        Assert.True(new FileInfo(fixturePath).Length > 0);
    }

    [Fact]
    [Trait("Category", "WriteForPoi")]
    public void Write_DocmWithParagraphsAndVba_CreatesFixtureForPoi()
    {
        var fixturePath = GetFixturePath("phase-docm-interop.docm");
        Directory.CreateDirectory(Path.GetDirectoryName(fixturePath)!);

        byte[] vbaBytes = ExtractZipEntry("example.docm", "word/vbaProject.bin");

        using var doc = new XWPFDocument();
        var para1 = doc.createParagraph();
        var run1 = para1.createRun();
        run1.setText("from dotnet-poi docm");
        run1.setBold(true);

        var para2 = doc.createParagraph();
        var run2 = para2.createRun();
        run2.setText("second paragraph");
        run2.setItalic(true);

        doc.setVBAProject(vbaBytes);

        using var stream = File.Create(fixturePath);
        doc.write(stream);

        Assert.True(File.Exists(fixturePath));
        Assert.True(new FileInfo(fixturePath).Length > 0);
    }

    [Fact]
    [Trait("Category", "WriteForPoi")]
    public void Write_PptmWithSlideAndVba_CreatesFixtureForPoi()
    {
        var fixturePath = GetFixturePath("phase-pptm-interop.pptm");
        Directory.CreateDirectory(Path.GetDirectoryName(fixturePath)!);

        byte[] vbaBytes = ExtractZipEntry("example.pptm", "ppt/vbaProject.bin");

        using var prs = new XMLSlideShow();
        var slide = prs.createSlide();
        var picIdx = prs.addPicture(LoadTestImage(), XSLFPictureData.PICTURE_TYPE_JPEG);
        var shape = prs.createPicture(slide, picIdx);
        shape.setAnchor(0, 0, XMLSlideShow.DefaultSlideCx, XMLSlideShow.DefaultSlideCy);
        prs.setVBAProject(vbaBytes);

        using var stream = File.Create(fixturePath);
        prs.write(stream);

        Assert.True(File.Exists(fixturePath));
        Assert.True(new FileInfo(fixturePath).Length > 0);
    }

    [Fact]
    [Trait("Category", "WriteForPoi")]
    public void Write_XlsmWithCellsAndVba_CreatesFixtureForPoi()
    {
        var fixturePath = GetFixturePath("phase-xlsm-interop.xlsm");
        Directory.CreateDirectory(Path.GetDirectoryName(fixturePath)!);

        // Extract VBA bytes from the test xlsm so the fixture carries a real VBA project.
        byte[] vbaBytes;
        using (var sourceStream = File.OpenRead("example.xlsm"))
        using (var sourceZip = new System.IO.Compression.ZipArchive(sourceStream, System.IO.Compression.ZipArchiveMode.Read))
        {
            var vbaEntry = sourceZip.GetEntry("xl/vbaProject.bin")
                ?? throw new InvalidDataException("example.xlsm has no xl/vbaProject.bin");
            using var vbaStream = vbaEntry.Open();
            using var ms = new MemoryStream();
            vbaStream.CopyTo(ms);
            vbaBytes = ms.ToArray();
        }

        using var workbook = new XSSFWorkbook();
        var sheet = workbook.createSheet("MacroSheet");
        var row = sheet.createRow(0);
        row.createCell(0).setCellValue("from dotnet-poi xlsm");
        row.createCell(1).setCellValue(99.5);
        workbook.setVBAProject(vbaBytes);

        using var stream = File.Create(fixturePath);
        workbook.write(stream);

        Assert.True(File.Exists(fixturePath));
        Assert.True(new FileInfo(fixturePath).Length > 0);
    }

    [Fact]
    [Trait("Category", "WriteForPoi")]
    public void Write_AutoFilterWorkbook_CreatesFixtureForPoi()
    {
        var fixturePath = GetFixturePath("phase-autofilter.xlsx");
        Directory.CreateDirectory(Path.GetDirectoryName(fixturePath)!);

        using var wb = new XSSFWorkbook();
        var sheet = wb.createSheet("Data");
        sheet.createRow(0).createCell(0).setCellValue("Category");
        sheet.createRow(0).createCell(1).setCellValue("Value");
        sheet.createRow(1).createCell(0).setCellValue("Food");
        sheet.createRow(1).createCell(1).setCellValue(100);
        sheet.createRow(2).createCell(0).setCellValue("Travel");
        sheet.createRow(2).createCell(1).setCellValue(200);
        sheet.setAutoFilter(new CellRangeAddress(0, 2, 0, 1));

        using var stream = File.Create(fixturePath);
        wb.write(stream);

        Assert.True(File.Exists(fixturePath));
        Assert.True(new FileInfo(fixturePath).Length > 0);
    }

    [Fact]
    [Trait("Category", "WriteForPoi")]
    public void Write_ProtectedWorkbook_CreatesFixtureForPoi()
    {
        var fixturePath = GetFixturePath("phase-protection.xlsx");
        Directory.CreateDirectory(Path.GetDirectoryName(fixturePath)!);

        using var wb = new XSSFWorkbook();
        var sheet = wb.createSheet("Data");
        sheet.createRow(0).createCell(0).setCellValue("protected cell");
        sheet.protectSheet(true);
        wb.protectWorkbook(true);

        using var stream = File.Create(fixturePath);
        wb.write(stream);

        Assert.True(File.Exists(fixturePath));
        Assert.True(new FileInfo(fixturePath).Length > 0);
    }

    [Fact]
    [Trait("Category", "WriteForPoi")]
    public void Write_ActiveSheetWorkbook_CreatesFixtureForPoi()
    {
        var fixturePath = GetFixturePath("phase-active-sheet.xlsx");
        Directory.CreateDirectory(Path.GetDirectoryName(fixturePath)!);

        using var wb = new XSSFWorkbook();
        wb.createSheet("First");
        wb.createSheet("Second");
        wb.createSheet("Third");
        wb.setActiveSheet(1);

        using var stream = File.Create(fixturePath);
        wb.write(stream);

        Assert.True(File.Exists(fixturePath));
        Assert.True(new FileInfo(fixturePath).Length > 0);
    }

    [Fact]
    [Trait("Category", "WriteForPoi")]
    public void Write_DocxWithFields_CreatesFixtureForPoi()
    {
        var fixturePath = GetFixturePath("phase-docx-fields.docx");
        Directory.CreateDirectory(Path.GetDirectoryName(fixturePath)!);

        using var doc = new XWPFDocument();

        // Paragraph with PAGE field
        var p1 = doc.createParagraph();
        p1.createRun().setText("Page ");
        p1.addField(" PAGE ");

        // Paragraph with TOC field
        var p2 = doc.createParagraph();
        p2.addField("TOC \\o \"1-3\" \\h \\z \\u");

        // Paragraph with text and a MERGEFIELD
        var p3 = doc.createParagraph();
        p3.createRun().setText("Hello ");
        p3.addField("MERGEFIELD CustomerName");

        using var stream = File.Create(fixturePath);
        doc.write(stream);

        Assert.True(File.Exists(fixturePath));
        Assert.True(new FileInfo(fixturePath).Length > 0);
    }

    private static byte[] ExtractZipEntry(string zipFileName, string entryName)
    {
        using var zip = new System.IO.Compression.ZipArchive(File.OpenRead(zipFileName), System.IO.Compression.ZipArchiveMode.Read);
        var entry = zip.GetEntry(entryName)
            ?? throw new InvalidDataException($"{zipFileName} has no entry {entryName}");
        using var s = entry.Open();
        using var ms = new MemoryStream();
        s.CopyTo(ms);
        return ms.ToArray();
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
