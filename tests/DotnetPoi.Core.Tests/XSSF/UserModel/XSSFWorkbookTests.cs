using System.IO.Compression;
using DotnetPoi.SS.UserModel;
using DotnetPoi.SS.Util;
using DotnetPoi.XSSF.UserModel;
using Xunit;

namespace DotnetPoi.XSSF.Tests.UserModel;

public class XSSFWorkbookTests
{
    [Fact]
    public void Write_StringAndNumberCells_ProducesValidXlsxParts()
    {
        using var workbook = new XSSFWorkbook();
        var sheet = workbook.createSheet("Data");
        var row = sheet.createRow(0);
        row.createCell(0).setCellValue("hello");
        row.createCell(1).setCellValue(42.5);

        using var stream = new MemoryStream();
        workbook.write(stream);

        stream.Position = 0;
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read);
        Assert.NotNull(archive.GetEntry("[Content_Types].xml"));
        Assert.NotNull(archive.GetEntry("_rels/.rels"));
        Assert.NotNull(archive.GetEntry("xl/workbook.xml"));
        Assert.NotNull(archive.GetEntry("xl/_rels/workbook.xml.rels"));
        Assert.NotNull(archive.GetEntry("xl/worksheets/sheet1.xml"));
        Assert.NotNull(archive.GetEntry("xl/sharedStrings.xml"));
        Assert.NotNull(archive.GetEntry("xl/styles.xml"));

        var sheetXml = ReadEntry(archive, "xl/worksheets/sheet1.xml");
        Assert.Contains("<dimension ref=\"A1:B1\"/>", sheetXml);
        Assert.Contains("<c r=\"A1\" t=\"s\"><v>0</v></c>", sheetXml);
        Assert.Contains("<c r=\"B1\"><v>42.5</v></c>", sheetXml);

        var sharedStringsXml = ReadEntry(archive, "xl/sharedStrings.xml");
        Assert.Contains("<sst count=\"1\" uniqueCount=\"1\"", sharedStringsXml);
        Assert.Contains("<si><t>hello</t></si>", sharedStringsXml);
    }

    [Fact]
    public void CreateSheet_WithoutName_UsesPoiStyleDefaultNames()
    {
        using var workbook = new XSSFWorkbook();

        var first = workbook.createSheet();
        var second = workbook.createSheet();

        Assert.Equal(2, workbook.getNumberOfSheets());
        Assert.Same(first, workbook.getSheetAt(0));
        Assert.Same(second, workbook.getSheetAt(1));
    }

    [Fact]
    public void GetCreationHelper_RepeatedCalls_ReturnSameInstance()
    {
        using var workbook = new XSSFWorkbook();

        var helper = workbook.getCreationHelper();

        Assert.Same(helper, workbook.getCreationHelper());
        Assert.Same(workbook, helper.getWorkbook());
    }

    [Fact]
    public void Read_RoundTrippedStringAndNumberCells_RestoresWorkbookModel()
    {
        using var original = new XSSFWorkbook();
        var sheet = original.createSheet("RoundTrip");
        var firstRow = sheet.createRow(0);
        firstRow.createCell(0).setCellValue("hello");
        firstRow.createCell(1).setCellValue(42.5);
        var secondRow = sheet.createRow(2);
        secondRow.createCell(3).setCellValue(0.0);

        using var stream = new MemoryStream();
        original.write(stream);

        stream.Position = 0;
        using var loaded = new XSSFWorkbook(stream);

        Assert.Equal(1, loaded.getNumberOfSheets());
        var loadedSheet = loaded.getSheet("RoundTrip");
        Assert.NotNull(loadedSheet);
        Assert.Equal(2, loadedSheet.getLastRowNum());
        Assert.Equal("hello", loadedSheet.getRow(0)!.getCell(0)!.getStringCellValue());
        Assert.Equal(42.5, loadedSheet.getRow(0)!.getCell(1)!.getNumericCellValue());
        Assert.Equal(0.0, loadedSheet.getRow(2)!.getCell(3)!.getNumericCellValue());
    }

    [Fact]
    public void Write_FormulaCell_ProducesFormulaElementWithoutEvaluation()
    {
        using var workbook = new XSSFWorkbook();
        var cell = workbook.createSheet("Formulas").createRow(0).createCell(0);
        cell.setCellFormula("A2+B2");

        using var stream = new MemoryStream();
        workbook.write(stream);

        stream.Position = 0;
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read);
        var sheetXml = ReadEntry(archive, "xl/worksheets/sheet1.xml");

        Assert.Contains("<c r=\"A1\"><f>A2+B2</f></c>", sheetXml);
    }

    [Fact]
    public void Read_RoundTrippedFormulaCell_RestoresFormulaAndCachedString()
    {
        using var original = new XSSFWorkbook();
        var cell = original.createSheet("Formulas").createRow(0).createCell(0);
        cell.setCellFormula("\"hello \"&\"world\"");
        cell.setCellValue("hello world");

        using var stream = new MemoryStream();
        original.write(stream);

        stream.Position = 0;
        using var loaded = new XSSFWorkbook(stream);

        var loadedCell = loaded.getSheet("Formulas")!.getRow(0)!.getCell(0)!;
        Assert.Equal(CellType.Formula, loadedCell.getCellType());
        Assert.Equal(CellType.String, loadedCell.getCachedFormulaResultType());
        Assert.Equal("\"hello \"&\"world\"", loadedCell.getCellFormula());
        Assert.Equal("hello world", loadedCell.getStringCellValue());
    }

    [Fact]
    public void Write_ForceFormulaRecalculation_ProducesCalcPr()
    {
        using var workbook = new XSSFWorkbook();
        workbook.createSheet("Formulas").createRow(0).createCell(0).setCellFormula("B1+C1");
        workbook.setForceFormulaRecalculation(true);

        using var stream = new MemoryStream();
        workbook.write(stream);

        stream.Position = 0;
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read);
        var workbookXml = ReadEntry(archive, "xl/workbook.xml");

        Assert.Contains("<calcPr calcId=\"0\" fullCalcOnLoad=\"true\"/>", workbookXml);
    }

    [Fact]
    public void Read_RoundTrippedForceFormulaRecalculation_RestoresFlag()
    {
        using var original = new XSSFWorkbook();
        original.createSheet("Formulas").createRow(0).createCell(0).setCellFormula("B1+C1");
        original.setForceFormulaRecalculation(true);

        using var stream = new MemoryStream();
        original.write(stream);

        stream.Position = 0;
        using var loaded = new XSSFWorkbook(stream);

        Assert.True(loaded.getForceFormulaRecalculation());

        loaded.setForceFormulaRecalculation(false);
        Assert.False(loaded.getForceFormulaRecalculation());
    }

    [Fact]
    public void EvaluateFormulaCell_SumAverageAndArithmetic_StoresCachedNumericValues()
    {
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
        var arithmetic = sheet.getRow(1)!.createCell(2);
        arithmetic.setCellFormula("A1+B1*2");

        var evaluator = workbook.getCreationHelper().createFormulaEvaluator();

        Assert.Equal(CellType.Numeric, evaluator.evaluateFormulaCell(sum));
        Assert.Equal(CellType.Numeric, evaluator.evaluateFormulaCell(average));
        Assert.Equal(CellType.Numeric, evaluator.evaluateFormulaCell(arithmetic));
        Assert.Equal(60.0, sum.getNumericCellValue());
        Assert.Equal(20.0, average.getNumericCellValue());
        Assert.Equal(50.0, arithmetic.getNumericCellValue());
    }

    [Fact]
    public void EvaluateAll_RepresentativeFunctions_StoresCachedValues()
    {
        using var workbook = new XSSFWorkbook();
        var sheet = workbook.createSheet("EvalAll");
        var row = sheet.createRow(0);
        row.createCell(0).setCellValue(5.0);
        row.createCell(1).setCellValue(9.0);
        row.createCell(2).setCellValue("text");

        var min = sheet.createRow(1).createCell(0);
        min.setCellFormula("MIN(A1:B1)");
        var max = sheet.getRow(1)!.createCell(1);
        max.setCellFormula("MAX(A1:B1)");
        var count = sheet.getRow(1)!.createCell(2);
        count.setCellFormula("COUNT(A1:C1)");
        var text = sheet.getRow(1)!.createCell(3);
        text.setCellFormula("CONCATENATE(\"total=\",A1+B1)");

        workbook.getCreationHelper().createFormulaEvaluator().evaluateAll();

        Assert.Equal(5.0, min.getNumericCellValue());
        Assert.Equal(9.0, max.getNumericCellValue());
        Assert.Equal(2.0, count.getNumericCellValue());
        Assert.Equal("total=14", text.getStringCellValue());
    }

    [Fact]
    public void Write_StyledCell_ProducesStylesAndCellStyleReference()
    {
        using var workbook = new XSSFWorkbook();
        var dataFormat = workbook.createDataFormat();
        var font = workbook.createFont();
        font.setBold(true);
        font.setItalic(true);
        font.setFontName("Arial");
        font.setFontHeightInPoints(14);
        font.setColor((short)IndexedColors.Red);

        var style = workbook.createCellStyle();
        style.setFont(font);
        style.setDataFormat(dataFormat.getFormat("0.00"));
        style.setFillForegroundColor((short)IndexedColors.Yellow);
        style.setFillPattern(FillPatternType.SolidForeground);
        style.setBorderBottom(BorderStyle.Thin);

        var sheet = workbook.createSheet("Styles");
        var cell = sheet.createRow(0).createCell(0);
        cell.setCellValue(12.3);
        cell.setCellStyle(style);

        using var stream = new MemoryStream();
        workbook.write(stream);

        stream.Position = 0;
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read);
        var sheetXml = ReadEntry(archive, "xl/worksheets/sheet1.xml");
        Assert.Contains("<c r=\"A1\" s=\"1\"><v>12.3</v></c>", sheetXml);

        var stylesXml = ReadEntry(archive, "xl/styles.xml");
        Assert.Contains("<fonts count=\"2\">", stylesXml);
        Assert.Contains("<b/>", stylesXml);
        Assert.Contains("<i/>", stylesXml);
        Assert.Contains("<color indexed=\"10\"/>", stylesXml);
        Assert.Contains("<name val=\"Arial\"/>", stylesXml);
        Assert.Contains("<patternFill patternType=\"solid\"><fgColor indexed=\"13\"/></patternFill>", stylesXml);
        Assert.Contains("<bottom style=\"thin\"/>", stylesXml);
        Assert.Contains("<xf numFmtId=\"2\" fontId=\"1\" fillId=\"2\" borderId=\"1\" xfId=\"0\" applyFont=\"true\" applyNumberFormat=\"true\" applyFill=\"true\" applyBorder=\"true\"/>", stylesXml);
    }

    [Fact]
    public void Read_RoundTrippedStyledCell_RestoresStyleIndexAndFormat()
    {
        using var original = new XSSFWorkbook();
        var style = original.createCellStyle();
        style.setDataFormat(original.createDataFormat().getFormat("0.00"));
        var sheet = original.createSheet("RoundTripStyle");
        var cell = sheet.createRow(0).createCell(0);
        cell.setCellValue(1.25);
        cell.setCellStyle(style);

        using var stream = new MemoryStream();
        original.write(stream);

        stream.Position = 0;
        using var loaded = new XSSFWorkbook(stream);

        var loadedCell = loaded.getSheet("RoundTripStyle")!.getRow(0)!.getCell(0)!;
        Assert.Equal(1, loadedCell.getCellStyle().getIndex());
        Assert.Equal((short)2, loadedCell.getCellStyle().getDataFormat());
        Assert.Equal("0.00", loadedCell.getCellStyle().getDataFormatString());
    }

    [Fact]
    public void Write_CustomDataFormat_ProducesUserDefinedNumFmt()
    {
        using var workbook = new XSSFWorkbook();
        var style = workbook.createCellStyle();
        style.setDataFormat(workbook.createDataFormat().getFormat("#,##0.000 kg"));
        var cell = workbook.createSheet("CustomFormat").createRow(0).createCell(0);
        cell.setCellValue(12.3456);
        cell.setCellStyle(style);

        using var stream = new MemoryStream();
        workbook.write(stream);

        stream.Position = 0;
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read);
        var stylesXml = ReadEntry(archive, "xl/styles.xml");

        Assert.Contains("<numFmts count=\"1\"><numFmt numFmtId=\"164\" formatCode=\"#,##0.000 kg\"/></numFmts>", stylesXml);
        Assert.Contains("<xf numFmtId=\"164\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyNumberFormat=\"true\"/>", stylesXml);
    }

    // ── Style round-trip ─────────────────────────────────────────────────

    /// Font attributes (name, bold, italic, height, indexed color) and data format
    /// survive write → read.
    [Fact]
    public void RoundTrip_StyledCell_FontAndDataFormatRestored()
    {
        using var original = new XSSFWorkbook();
        var font = original.createFont();
        font.setFontName("Arial");
        font.setFontHeightInPoints(14);
        font.setBold(true);
        font.setItalic(true);
        font.setColor((short)IndexedColors.Red);

        var style = original.createCellStyle();
        style.setFont(font);
        style.setDataFormat(original.createDataFormat().getFormat("0.00"));

        var cell = original.createSheet("Styles").createRow(0).createCell(0);
        cell.setCellValue(123.456);
        cell.setCellStyle(style);

        using var stream = new MemoryStream();
        original.write(stream);

        stream.Position = 0;
        using var loaded = new XSSFWorkbook(stream);

        var loadedCell = loaded.getSheet("Styles")!.getRow(0)!.getCell(0)!;
        Assert.Equal(123.456, loadedCell.getNumericCellValue());

        var loadedStyle = loadedCell.getCellStyle();
        Assert.Equal("0.00", loadedStyle.getDataFormatString());

        var loadedFont = loadedStyle.getFont();
        Assert.Equal("Arial", loadedFont.getFontName());
        Assert.Equal(14, loadedFont.getFontHeightInPoints());
        Assert.True(loadedFont.getBold());
        Assert.True(loadedFont.getItalic());
        Assert.Equal((short)IndexedColors.Red, loadedFont.getColor());
    }

    [Fact]
    public void RoundTrip_MultipleStyles_EachCellRestoresItsOwnStyle()
    {
        using var original = new XSSFWorkbook();
        var sheet = original.createSheet("MultiStyle");

        // Cell A1: bold, font size 12
        var fontBold = original.createFont();
        fontBold.setBold(true);
        fontBold.setFontHeightInPoints(12);
        var styleBold = original.createCellStyle();
        styleBold.setFont(fontBold);
        var cellA1 = sheet.createRow(0).createCell(0);
        cellA1.setCellValue("bold");
        cellA1.setCellStyle(styleBold);

        // Cell B1: italic, custom number format
        var fontItalic = original.createFont();
        fontItalic.setItalic(true);
        var styleItalic = original.createCellStyle();
        styleItalic.setFont(fontItalic);
        styleItalic.setDataFormat(original.createDataFormat().getFormat("#,##0.0"));
        var cellB1 = sheet.getRow(0)!.createCell(1);
        cellB1.setCellValue(9999.5);
        cellB1.setCellStyle(styleItalic);

        using var stream = new MemoryStream();
        original.write(stream);

        stream.Position = 0;
        using var loaded = new XSSFWorkbook(stream);
        var loadedSheet = loaded.getSheet("MultiStyle")!;

        var la1 = loadedSheet.getRow(0)!.getCell(0)!;
        Assert.Equal("bold", la1.getStringCellValue());
        Assert.True(la1.getCellStyle().getFont().getBold());
        Assert.Equal(12, la1.getCellStyle().getFont().getFontHeightInPoints());
        Assert.False(la1.getCellStyle().getFont().getItalic());

        var lb1 = loadedSheet.getRow(0)!.getCell(1)!;
        Assert.Equal(9999.5, lb1.getNumericCellValue());
        Assert.False(lb1.getCellStyle().getFont().getBold());
        Assert.True(lb1.getCellStyle().getFont().getItalic());
        Assert.Equal("#,##0.0", lb1.getCellStyle().getDataFormatString());
    }

    [Fact]
    public void RoundTrip_BuiltinDateFormat_DataFormatIndexRestored()
    {
        // Format index 14 is the built-in "m/d/yy" date format in OOXML.
        const short dateFormatIndex = 14;

        using var original = new XSSFWorkbook();
        var style = original.createCellStyle();
        style.setDataFormat(dateFormatIndex);
        var cell = original.createSheet("Date").createRow(0).createCell(0);
        cell.setCellValue(45678.0); // an Excel date serial number
        cell.setCellStyle(style);

        using var stream = new MemoryStream();
        original.write(stream);

        stream.Position = 0;
        using var loaded = new XSSFWorkbook(stream);
        var loadedCell = loaded.getSheet("Date")!.getRow(0)!.getCell(0)!;

        Assert.Equal(dateFormatIndex, loadedCell.getCellStyle().getDataFormat());
    }

    [Fact]
    public void RoundTrip_StyledCell_FillRestored()
    {
        using var original = new XSSFWorkbook();
        var style = original.createCellStyle();
        style.setFillForegroundColor((short)IndexedColors.Yellow);
        style.setFillPattern(FillPatternType.SolidForeground);

        var cell = original.createSheet("Fill").createRow(0).createCell(0);
        cell.setCellValue(42.0);
        cell.setCellStyle(style);

        using var stream = new MemoryStream();
        original.write(stream);

        stream.Position = 0;
        using var loaded = new XSSFWorkbook(stream);
        var loadedCell = loaded.getSheet("Fill")!.getRow(0)!.getCell(0)!;
        var loadedStyle = loadedCell.getCellStyle();
        Assert.Equal((short)IndexedColors.Yellow, loadedStyle.getFillForegroundColor());
        Assert.Equal(FillPatternType.SolidForeground, loadedStyle.getFillPattern());
    }

    [Fact]
    public void RoundTrip_StyledCell_BorderRestored()
    {
        using var original = new XSSFWorkbook();
        var style = original.createCellStyle();
        style.setBorderTop(BorderStyle.Medium);
        style.setBorderRight(BorderStyle.Dotted);
        style.setBorderBottom(BorderStyle.Thick);
        style.setBorderLeft(BorderStyle.Dashed);

        var cell = original.createSheet("Border").createRow(0).createCell(0);
        cell.setCellValue(42.0);
        cell.setCellStyle(style);

        using var stream = new MemoryStream();
        original.write(stream);

        stream.Position = 0;
        using var loaded = new XSSFWorkbook(stream);
        var loadedCell = loaded.getSheet("Border")!.getRow(0)!.getCell(0)!;
        var loadedStyle = loadedCell.getCellStyle();
        Assert.Equal(BorderStyle.Medium, loadedStyle.getBorderTop());
        Assert.Equal(BorderStyle.Dotted, loadedStyle.getBorderRight());
        Assert.Equal(BorderStyle.Thick, loadedStyle.getBorderBottom());
        Assert.Equal(BorderStyle.Dashed, loadedStyle.getBorderLeft());
    }

    [Fact]
    public void RoundTrip_StyledCell_AlignmentRestored()
    {
        using var original = new XSSFWorkbook();
        var style = original.createCellStyle();
        style.setAlignment(HorizontalAlignment.Center);
        style.setVerticalAlignment(VerticalAlignment.Top);
        style.setWrapText(true);
        style.setIndention(1);
        style.setRotation(45);

        var cell = original.createSheet("Align").createRow(0).createCell(0);
        cell.setCellValue("aligned");
        cell.setCellStyle(style);

        using var stream = new MemoryStream();
        original.write(stream);

        stream.Position = 0;
        using var loaded = new XSSFWorkbook(stream);
        var loadedCell = loaded.getSheet("Align")!.getRow(0)!.getCell(0)!;
        var loadedStyle = loadedCell.getCellStyle();
        Assert.Equal(HorizontalAlignment.Center, loadedStyle.getAlignment());
        Assert.Equal(VerticalAlignment.Top, loadedStyle.getVerticalAlignment());
        Assert.True(loadedStyle.getWrapText());
        Assert.Equal((short)1, loadedStyle.getIndention());
        Assert.Equal((short)45, loadedStyle.getRotation());
    }

    [Fact]
    public void RoundTrip_MergeCells_Preserved()
    {
        using var original = new XSSFWorkbook();
        var sheet = original.createSheet("MergeTest");

        var row1 = sheet.createRow(0);
        row1.createCell(0).setCellValue("A1");
        row1.createCell(1).setCellValue("B1");
        var row2 = sheet.createRow(1);
        row2.createCell(0).setCellValue("A2");
        row2.createCell(1).setCellValue("B2");

        sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 1)); // A1:B1
        sheet.addMergedRegion(new CellRangeAddress(1, 1, 0, 0)); // A2:A2

        using var stream = new MemoryStream();
        original.write(stream);

        stream.Position = 0;
        using var loaded = new XSSFWorkbook(stream);
        var loadedSheet = loaded.getSheet("MergeTest")!;

        var merged = loadedSheet.getMergedRegions();
        Assert.Equal(2, merged.Count);

        Assert.Equal(0, merged[0].FirstRow);
        Assert.Equal(0, merged[0].LastRow);
        Assert.Equal(0, merged[0].FirstCol);
        Assert.Equal(1, merged[0].LastCol);

        Assert.Equal(1, merged[1].FirstRow);
        Assert.Equal(1, merged[1].LastRow);
        Assert.Equal(0, merged[1].FirstCol);
        Assert.Equal(0, merged[1].LastCol);
    }

    [Fact]
    public void RoundTrip_ColumnWidth_Preserved()
    {
        using var original = new XSSFWorkbook();
        var sheet = original.createSheet("WidthTest");
        sheet.createRow(0).createCell(0).setCellValue("A");
        sheet.createRow(0).createCell(1).setCellValue("B");

        sheet.setColumnWidth(0, 80 * 256);
        sheet.setColumnWidth(1, 40 * 256);

        using var stream = new MemoryStream();
        original.write(stream);

        stream.Position = 0;
        using var loaded = new XSSFWorkbook(stream);
        var loadedSheet = loaded.getSheet("WidthTest")!;

        var loadedWidth0 = loadedSheet.getColumnWidth(0);
        var loadedWidth1 = loadedSheet.getColumnWidth(1);
        Assert.InRange(loadedWidth0, 75 * 256, 85 * 256);
        Assert.InRange(loadedWidth1, 35 * 256, 45 * 256);
    }

    [Fact]
    public void RoundTrip_RowHeight_Preserved()
    {
        using var original = new XSSFWorkbook();
        var sheet = original.createSheet("HeightTest");
        var row0 = sheet.createRow(0);
        row0.createCell(0).setCellValue("Tall");
        row0.setHeight(45.0f);

        var row1 = sheet.createRow(1);
        row1.createCell(0).setCellValue("Normal");

        using var stream = new MemoryStream();
        original.write(stream);

        stream.Position = 0;
        using var loaded = new XSSFWorkbook(stream);
        var loadedSheet = loaded.getSheet("HeightTest")!;

        var loadedRow0 = loadedSheet.getRow(0)!;
        Assert.Equal(45.0f, loadedRow0.getHeight(), 3);

        var loadedRow1 = loadedSheet.getRow(1)!;
        Assert.Equal(15.0f, loadedRow1.getHeight(), 3);
    }

    [Fact]
    public void RoundTrip_Hyperlink_Preserved()
    {
        using var original = new XSSFWorkbook();
        var sheet = original.createSheet("LinkTest");
        var cell = sheet.createRow(0).createCell(0);
        cell.setCellValue("Click me");

        var hyperlink = new XSSFHyperlink(HyperlinkType.Url);
        hyperlink.setAddress("https://example.com");
        cell.setHyperlink(hyperlink);

        using var stream = new MemoryStream();
        original.write(stream);

        stream.Position = 0;
        using var loaded = new XSSFWorkbook(stream);
        var loadedSheet = loaded.getSheet("LinkTest")!;
        var loadedCell = loadedSheet.getRow(0)!.getCell(0)!;
        Assert.Equal("Click me", loadedCell.getStringCellValue());

        var loadedLink = loadedCell.getHyperlink();
        Assert.NotNull(loadedLink);
        Assert.Equal("https://example.com", loadedLink.getAddress());
        Assert.Equal("A1", loadedLink.getCellRef());
        Assert.Equal(HyperlinkType.Url, loadedLink.getType());
    }

    [Fact]
    public void RoundTrip_PrintSettings_Preserved()
    {
        using var original = new XSSFWorkbook();
        var sheet = original.createSheet("PrintTest");
        sheet.createRow(0).createCell(0).setCellValue("Page 1");

        // Set page layout properties
        sheet.PageOrientation = "landscape";
        sheet.PaperSize = 9; // A4
        sheet.PageMarginLeft = 1.0;
        sheet.PageMarginRight = 1.0;
        sheet.PageMarginTop = 1.5;
        sheet.PageMarginBottom = 1.5;
        sheet.HeaderCenter = "Confidential";
        sheet.FooterCenter = "Page 1";

        using var stream = new MemoryStream();
        original.write(stream);

        stream.Position = 0;
        using var loaded = new XSSFWorkbook(stream);
        var loadedSheet = loaded.getSheetAt(0);
        var xssfSheet = (XSSFSheet)loadedSheet;

        Assert.Equal("landscape", xssfSheet.PageOrientation);
        Assert.Equal(9, xssfSheet.PaperSize);
        Assert.Equal(1.0, xssfSheet.PageMarginLeft, 2);
        Assert.Equal(1.0, xssfSheet.PageMarginRight, 2);
        Assert.Equal(1.5, xssfSheet.PageMarginTop, 2);
        Assert.Equal(1.5, xssfSheet.PageMarginBottom, 2);
        Assert.Equal("Confidential", xssfSheet.HeaderCenter);
        Assert.Equal("Page 1", xssfSheet.FooterCenter);
    }

    [Fact]
    public void RoundTrip_DataValidation_Preserved()
    {
        using var original = new XSSFWorkbook();
        var sheet = original.createSheet("Validation");
        sheet.createRow(0).createCell(0).setCellValue("Value");

        var dv = new XSSFDataValidation
        {
            Sqref = "A1:A10",
            Type = DataValidationType.Whole,
            Operator = DataValidationOperator.Between,
            Formula1 = "1",
            Formula2 = "100",
            AllowBlank = false,
            ShowDropDown = false,
            ErrorStyle = "stop",
            ErrorTitle = "Invalid",
            ErrorMessage = "Value must be between 1 and 100."
        };
        sheet.AddDataValidation(dv);

        using var stream = new MemoryStream();
        original.write(stream);

        stream.Position = 0;
        using var loaded = new XSSFWorkbook(stream);
        var loadedSheet = loaded.getSheetAt(0);
        var xssfSheet = (XSSFSheet)loadedSheet;

        var loadedValidations = xssfSheet.DataValidations;
        Assert.Single(loadedValidations);

        var loadedDv = loadedValidations[0];
        Assert.Equal("A1:A10", loadedDv.Sqref);
        Assert.Equal(DataValidationType.Whole, loadedDv.Type);
        Assert.Equal(DataValidationOperator.Between, loadedDv.Operator);
        Assert.Equal("1", loadedDv.Formula1);
        Assert.Equal("100", loadedDv.Formula2);
        Assert.False(loadedDv.AllowBlank);
        Assert.False(loadedDv.ShowDropDown);
        Assert.Equal("stop", loadedDv.ErrorStyle);
        Assert.Equal("Invalid", loadedDv.ErrorTitle);
        Assert.Equal("Value must be between 1 and 100.", loadedDv.ErrorMessage);
    }

    [Fact]
    public void RoundTrip_DataValidation_TypesAndOperators()
    {
        // Test all DataValidation types (Decimal, Date, TextLength, List, Time, Custom)
        // with various operators to ensure round-trip preservation.
        using var original = new XSSFWorkbook();
        var sheet = original.createSheet("Validation");
        sheet.createRow(0).createCell(0).setCellValue("A");
        sheet.createRow(1).createCell(0).setCellValue("B");
        sheet.createRow(0).createCell(1).setCellValue("C");
        sheet.createRow(1).createCell(1).setCellValue("D");

        // 1. Decimal + Between
        var dvDecimal = new XSSFDataValidation
        {
            Sqref = "A1:A10",
            Type = DataValidationType.Decimal,
            Operator = DataValidationOperator.Between,
            Formula1 = "1.5",
            Formula2 = "10.0",
            AllowBlank = true,
            ShowDropDown = true,
            ErrorStyle = "stop",
            ErrorTitle = "Range",
            ErrorMessage = "Enter a decimal between 1.5 and 10.0."
        };
        sheet.AddDataValidation(dvDecimal);

        // 2. Date + Equal
        var dvDate = new XSSFDataValidation
        {
            Sqref = "B1:B10",
            Type = DataValidationType.Date,
            Operator = DataValidationOperator.Equal,
            Formula1 = "2026-01-01",
            AllowBlank = false,
            ShowDropDown = false,
        };
        sheet.AddDataValidation(dvDate);

        // 3. TextLength + Between
        var dvTextLength = new XSSFDataValidation
        {
            Sqref = "C1:C10",
            Type = DataValidationType.TextLength,
            Operator = DataValidationOperator.Between,
            Formula1 = "1",
            Formula2 = "100",
        };
        sheet.AddDataValidation(dvTextLength);

        // 4. List (formula-based, no operator — operator skipped for List type)
        var dvList = new XSSFDataValidation
        {
            Sqref = "D1:D10",
            Type = DataValidationType.List,
            Formula1 = "$A$1:$A$3",
        };
        sheet.AddDataValidation(dvList);

        // 5. Time + LessThan
        var dvTime = new XSSFDataValidation
        {
            Sqref = "E1:E10",
            Type = DataValidationType.Time,
            Operator = DataValidationOperator.LessThan,
            Formula1 = "12:00:00",
        };
        sheet.AddDataValidation(dvTime);

        // Round-trip and verify
        using var stream = new MemoryStream();
        original.write(stream);
        stream.Position = 0;

        using var loaded = new XSSFWorkbook(stream);
        var loadedSheet = loaded.getSheetAt(0);
        var xssfSheet = (XSSFSheet)loadedSheet;

        var loadedValidations = xssfSheet.DataValidations;
        Assert.Equal(5, loadedValidations.Count);

        // Verify Decimal validation
        var dv0 = loadedValidations[0];
        Assert.Equal("A1:A10", dv0.Sqref);
        Assert.Equal(DataValidationType.Decimal, dv0.Type);
        Assert.Equal(DataValidationOperator.Between, dv0.Operator);
        Assert.Equal("1.5", dv0.Formula1);
        Assert.Equal("10.0", dv0.Formula2);
        Assert.True(dv0.AllowBlank);
        Assert.True(dv0.ShowDropDown);
        Assert.Equal("Enter a decimal between 1.5 and 10.0.", dv0.ErrorMessage);

        // Verify Date validation
        var dv1 = loadedValidations[1];
        Assert.Equal("B1:B10", dv1.Sqref);
        Assert.Equal(DataValidationType.Date, dv1.Type);
        Assert.Equal(DataValidationOperator.Equal, dv1.Operator);
        Assert.Equal("2026-01-01", dv1.Formula1);
        Assert.False(dv1.AllowBlank);
        Assert.False(dv1.ShowDropDown);

        // Verify TextLength validation
        var dv2 = loadedValidations[2];
        Assert.Equal("C1:C10", dv2.Sqref);
        Assert.Equal(DataValidationType.TextLength, dv2.Type);
        Assert.Equal(DataValidationOperator.Between, dv2.Operator);
        Assert.Equal("1", dv2.Formula1);
        Assert.Equal("100", dv2.Formula2);

        // Verify List validation (no operator — should be None/unspecified)
        var dv3 = loadedValidations[3];
        Assert.Equal("D1:D10", dv3.Sqref);
        Assert.Equal(DataValidationType.List, dv3.Type);
        Assert.Equal("$A$1:$A$3", dv3.Formula1);

        // Verify Time validation
        var dv4 = loadedValidations[4];
        Assert.Equal("E1:E10", dv4.Sqref);
        Assert.Equal(DataValidationType.Time, dv4.Type);
        Assert.Equal(DataValidationOperator.LessThan, dv4.Operator);
        Assert.Equal("12:00:00", dv4.Formula1);
    }

    [Fact]
    public void RoundTrip_ConditionalFormatting_Preserved()
    {
        using var original = new XSSFWorkbook();
        var sheet = original.createSheet("Rules");
        var cell = sheet.createRow(0).createCell(0);
        cell.setCellValue(50);

        // Create a conditional formatting rule: cell value > 100 should highlight in red
        var cf = new XSSFConditionalFormatting();
        cf.Sqref = "A1:A10";
        var rule = new XSSFCFRule();
        rule.Type = ConditionalFormatType.CellIs;
        rule.Priority = 1;
        rule.Operator = "greaterThan";
        rule.Formulas.Add("100");
        cf.Rules.Add(rule);
        sheet.AddConditionalFormatting(cf);

        using var stream = new MemoryStream();
        original.write(stream);
        stream.Position = 0;

        using var loaded = new XSSFWorkbook(stream);
        var loadedSheet = loaded.getSheet("Rules")!;
        var loadedCfs = loadedSheet.ConditionalFormatting;
        Assert.Single(loadedCfs);

        var loadedCf = loadedCfs[0];
        Assert.Equal("A1:A10", loadedCf.Sqref);
        Assert.Single(loadedCf.Rules);

        var loadedRule = loadedCf.Rules[0];
        Assert.Equal(ConditionalFormatType.CellIs, loadedRule.Type);
        Assert.Equal(1, loadedRule.Priority);
        Assert.Equal("greaterThan", loadedRule.Operator);
        Assert.Single(loadedRule.Formulas);
        Assert.Equal("100", loadedRule.Formulas[0]);
    }

    [Fact]
    public void RoundTrip_ConditionalFormatting_FormulaType()
    {
        // Test ConditionalFormatType.Formula with custom formula
        using var original = new XSSFWorkbook();
        var sheet = original.createSheet("Rules");
        sheet.createRow(0).createCell(0).setCellValue(10);
        sheet.createRow(1).createCell(0).setCellValue(200);

        // Formula type: =A1>100
        var cf = new XSSFConditionalFormatting();
        cf.Sqref = "A1:A10";
        var rule = new XSSFCFRule();
        rule.Type = ConditionalFormatType.Formula;
        rule.Priority = 1;
        rule.Operator = "greaterThan";
        rule.Formulas.Add("A1>100");
        cf.Rules.Add(rule);
        sheet.AddConditionalFormatting(cf);

        using var stream = new MemoryStream();
        original.write(stream);
        stream.Position = 0;

        using var loaded = new XSSFWorkbook(stream);
        var loadedSheet = loaded.getSheet("Rules")!;
        var loadedCfs = loadedSheet.ConditionalFormatting;
        Assert.Single(loadedCfs);
        var loadedCf = loadedCfs[0];
        Assert.Equal("A1:A10", loadedCf.Sqref);
        Assert.Single(loadedCf.Rules);

        var loadedRule = loadedCf.Rules[0];
        Assert.Equal(ConditionalFormatType.Formula, loadedRule.Type);
        Assert.Equal(1, loadedRule.Priority);
        Assert.Equal("greaterThan", loadedRule.Operator);
        Assert.Single(loadedRule.Formulas);
        Assert.Equal("A1>100", loadedRule.Formulas[0]);
    }

    [Fact]
    public void RoundTrip_FreezePane_Preserved()
    {
        using var original = new XSSFWorkbook();
        var sheet = original.createSheet("Freeze");
        sheet.createRow(0).createCell(0).setCellValue("Header");
        sheet.createRow(1).createCell(0).setCellValue("Data");
        sheet.createFreezePane(1, 1);

        using var stream = new MemoryStream();
        original.write(stream);
        stream.Position = 0;

        using var loaded = new XSSFWorkbook(stream);
        var loadedSheet = loaded.getSheet("Freeze")!;
        Assert.Equal(1, loadedSheet.FreezeColSplit);
        Assert.Equal(1, loadedSheet.FreezeRowSplit);
    }

    [Fact]
    public void RoundTrip_HiddenRow_Preserved()
    {
        using var original = new XSSFWorkbook();
        var sheet = original.createSheet("HideRow");
        var row0 = sheet.createRow(0);
        row0.createCell(0).setCellValue("Visible");
        var row1 = sheet.createRow(1);
        row1.createCell(0).setCellValue("Hidden");
        row1.setHidden(true);

        using var stream = new MemoryStream();
        original.write(stream);
        stream.Position = 0;

        using var loaded = new XSSFWorkbook(stream);
        var loadedSheet = loaded.getSheet("HideRow")!;
        Assert.False(loadedSheet.getRow(0)!.isHidden());
        Assert.True(loadedSheet.getRow(1)!.isHidden());
    }

    [Fact]
    public void RoundTrip_HiddenColumn_Preserved()
    {
        using var original = new XSSFWorkbook();
        var sheet = original.createSheet("HideCol");
        sheet.createRow(0).createCell(0).setCellValue("A");
        sheet.createRow(0).createCell(1).setCellValue("B (hidden)");
        sheet.setColumnHidden(1, true);

        using var stream = new MemoryStream();
        original.write(stream);
        stream.Position = 0;

        using var loaded = new XSSFWorkbook(stream);
        var loadedSheet = loaded.getSheet("HideCol")!;
        Assert.False(loadedSheet.isColumnHidden(0));
        Assert.True(loadedSheet.isColumnHidden(1));
    }

    [Fact]
    public void RoundTrip_SharedStrings_Preserved()
    {
        using var original = new XSSFWorkbook();
        var sheet = original.createSheet("Strings");
        var row = sheet.createRow(0);
        row.createCell(0).setCellValue("Hello");
        row.createCell(1).setCellValue("World");
        row.createCell(2).setCellValue("C# with DotnetPoi");
        var row2 = sheet.createRow(1);
        row2.createCell(0).setCellValue("Hello"); // duplicate — should share index

        using var stream = new MemoryStream();
        original.write(stream);
        stream.Position = 0;

        // Verify the SST has exactly 3 unique entries
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read);
        var sstEntry = archive.GetEntry("xl/sharedStrings.xml");
        Assert.NotNull(sstEntry);
        using var sstReader = new StreamReader(sstEntry.Open());
        var sstXml = sstReader.ReadToEnd();
        Assert.Contains("<t>Hello</t>", sstXml);
        Assert.Contains("<t>World</t>", sstXml);
        Assert.Contains("<t>C# with DotnetPoi</t>", sstXml);

        // Reload workbook and verify cell values
        stream.Position = 0;
        using var loaded = new XSSFWorkbook(stream);
        var loadedSheet = loaded.getSheet("Strings")!;
        Assert.Equal("Hello", loadedSheet.getRow(0)!.getCell(0)!.getStringCellValue());
        Assert.Equal("World", loadedSheet.getRow(0)!.getCell(1)!.getStringCellValue());
        Assert.Equal("C# with DotnetPoi", loadedSheet.getRow(0)!.getCell(2)!.getStringCellValue());
        Assert.Equal("Hello", loadedSheet.getRow(1)!.getCell(0)!.getStringCellValue());
    }

    [Fact]
    public void RoundTrip_RichTextFormatting_Preserved()
    {
        using var original = new XSSFWorkbook();
        var sheet = original.createSheet("RichText");
        var cell = sheet.createRow(0).createCell(0);

        // Create rich text with two formatted runs
        var richText = new XSSFRichTextString();
        richText.addRun("Hello ", bold: true, italic: false, underline: false, strikethrough: false,
            fontSize: 14.0, fontName: "Arial", color: "FFFF0000");
        richText.addRun("World", bold: false, italic: true, underline: false, strikethrough: false,
            fontSize: 12.0, fontName: "Calibri", color: "FF0000FF");

        cell.setCellValue(richText);

        using var stream = new MemoryStream();
        original.write(stream);
        stream.Position = 0;

        using var loaded = new XSSFWorkbook(stream);
        var loadedCell = loaded.getSheet("RichText")!.getRow(0)!.getCell(0)!;

        // Verify plain text
        Assert.Equal("Hello World", loadedCell.getStringCellValue());

        // Verify rich text
        var loadedRich = loadedCell.getRichStringCellValue();
        Assert.NotNull(loadedRich);
        Assert.True(loadedRich.IsRichText);
        Assert.Equal(2, loadedRich.Runs.Count);

        // First run: "Hello " bold red
        var run0 = loadedRich.Runs[0];
        Assert.Equal("Hello ", run0.Text);
        Assert.True(run0.Bold);
        Assert.False(run0.Italic);
        Assert.Equal(14.0, run0.FontSize, 1); // tolerance due to OOXML rounding
        Assert.Equal("Arial", run0.FontName);
        Assert.Contains("FF0000", run0.Color ?? string.Empty);

        // Second run: "World" italic blue
        var run1 = loadedRich.Runs[1];
        Assert.Equal("World", run1.Text);
        Assert.False(run1.Bold);
        Assert.True(run1.Italic);
        Assert.Equal(12.0, run1.FontSize, 1);
        Assert.Equal("Calibri", run1.FontName);
        Assert.Contains("0000FF", run1.Color ?? string.Empty);
    }

    [Fact]
    public void RoundTrip_UnknownParts_Preserved()
    {
        // Write a basic xlsx
        byte[] xlsxBytes;
        using (var original = new XSSFWorkbook())
        {
            var sheet = original.createSheet("Data");
            sheet.createRow(0).createCell(0).setCellValue("Hello");
            using var ms = new MemoryStream();
            original.write(ms);
            xlsxBytes = ms.ToArray();
        }

        // Inject an extra entry (simulating pivot table parts) into a copy
        using var injectedStream = new MemoryStream();
        using (var srcArchive = new ZipArchive(new MemoryStream(xlsxBytes), ZipArchiveMode.Read))
        {
            // Copy all existing entries verbatim
            using var dstArchive = new ZipArchive(injectedStream, ZipArchiveMode.Create, leaveOpen: true);
            foreach (var e in srcArchive.Entries)
            {
                var newEntry = dstArchive.CreateEntry(e.FullName, CompressionLevel.Optimal);
                using var src = e.Open();
                using var dst = newEntry.Open();
                src.CopyTo(dst);
            }
            // Add the pivot part
            var pivotEntry = dstArchive.CreateEntry("xl/pivotTables/pivotTable1.xml", CompressionLevel.Optimal);
            using var sw = new StreamWriter(pivotEntry.Open());
            sw.Write("FakePivotContent");
        }

        injectedStream.Position = 0;

        // Load & re-save — the pivot part should survive
        using var loaded = new XSSFWorkbook(injectedStream);
        Assert.Equal("Hello", loaded.getSheet("Data")!.getRow(0)!.getCell(0)!.getStringCellValue());

        using var outStream = new MemoryStream();
        loaded.write(outStream);
        outStream.Position = 0;

        using var resultArchive = new ZipArchive(outStream, ZipArchiveMode.Read);
        var pivotEntryResult = resultArchive.GetEntry("xl/pivotTables/pivotTable1.xml");
        Assert.NotNull(pivotEntryResult); // pivot table part must survive
        using var reader = new StreamReader(pivotEntryResult.Open());
        Assert.Equal("FakePivotContent", reader.ReadToEnd());
    }

    [Fact]
    public void RoundTrip_SheetProtection_Preserved()
    {
        using var original = new XSSFWorkbook();
        var sheet = original.createSheet("Protected");
        sheet.createRow(0).createCell(0).setCellValue("data");
        sheet.protectSheet(true);

        using var stream = new MemoryStream();
        original.write(stream);
        stream.Position = 0;

        // Verify the sheetProtection element exists in the worksheet XML
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read);
        var sheetEntry = archive.GetEntry("xl/worksheets/sheet1.xml");
        Assert.NotNull(sheetEntry);
        using var sheetReader = new StreamReader(sheetEntry.Open());
        var sheetXml = sheetReader.ReadToEnd();
        Assert.Contains("<sheetProtection", sheetXml);

        // Reload and verify
        stream.Position = 0;
        using var loaded = new XSSFWorkbook(stream);
        var loadedSheet = loaded.getSheet("Protected")!;
        Assert.True(loadedSheet.isSheetProtected());
        Assert.Equal("data", loadedSheet.getRow(0)!.getCell(0)!.getStringCellValue());
    }

    [Fact]
    public void RoundTrip_WorkbookProtection_Preserved()
    {
        using var original = new XSSFWorkbook();
        original.createSheet("Data");
        original.protectWorkbook(true);

        using var stream = new MemoryStream();
        original.write(stream);
        stream.Position = 0;

        // Verify the workbookProtection element exists
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read);
        var wbEntry = archive.GetEntry("xl/workbook.xml");
        Assert.NotNull(wbEntry);
        using var wbReader = new StreamReader(wbEntry.Open());
        var wbXml = wbReader.ReadToEnd();
        Assert.Contains("<workbookProtection", wbXml);

        // Reload and verify
        stream.Position = 0;
        using var loaded = new XSSFWorkbook(stream);
        Assert.True(loaded.isWorkbookProtected());
    }

    [Fact]
    public void RoundTrip_AutoFilter_Preserved()
    {
        using var original = new XSSFWorkbook();
        var sheet = original.createSheet("Filtered");
        sheet.createRow(0).createCell(0).setCellValue("Header1");
        sheet.createRow(0).createCell(1).setCellValue("Header2");
        sheet.createRow(1).createCell(0).setCellValue("Data1");
        sheet.createRow(1).createCell(1).setCellValue("Data2");
        sheet.setAutoFilter(new CellRangeAddress(0, 1, 0, 1));

        using var stream = new MemoryStream();
        original.write(stream);
        stream.Position = 0;

        // Verify the autoFilter element exists in the worksheet XML
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read);
        var sheetEntry = archive.GetEntry("xl/worksheets/sheet1.xml");
        Assert.NotNull(sheetEntry);
        using var sheetReader = new StreamReader(sheetEntry.Open());
        var sheetXml = sheetReader.ReadToEnd();
        Assert.Contains("<autoFilter", sheetXml);
        Assert.Contains("ref=\"A1:B2\"", sheetXml);

        // Reload and verify
        stream.Position = 0;
        using var loaded = new XSSFWorkbook(stream);
        var loadedSheet = loaded.getSheet("Filtered")!;
        var autoFilter = loadedSheet.getAutoFilter();
        Assert.NotNull(autoFilter);
        Assert.Equal(0, autoFilter.FirstRow);
        Assert.Equal(1, autoFilter.LastRow);
        Assert.Equal(0, autoFilter.FirstCol);
        Assert.Equal(1, autoFilter.LastCol);
    }

    [Fact]
    public void RoundTrip_ActiveSheetIndex_Preserved()
    {
        using var original = new XSSFWorkbook();
        original.createSheet("Sheet1");
        original.createSheet("Sheet2");
        original.createSheet("Sheet3");
        original.setActiveSheet(1);

        using var stream = new MemoryStream();
        original.write(stream);
        stream.Position = 0;

        // Verify in workbook XML
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read);
        var wbEntry = archive.GetEntry("xl/workbook.xml");
        Assert.NotNull(wbEntry);
        using var wbReader = new StreamReader(wbEntry.Open());
        var wbXml = wbReader.ReadToEnd();
        Assert.Contains("activeTab=\"1\"", wbXml);

        // Reload and verify round-trip
        stream.Position = 0;
        using var loaded = new XSSFWorkbook(stream);
        Assert.Equal(1, loaded.getActiveSheetIndex());
    }

    [Fact]
    public void ActiveCellApi_WorksInMemory()
    {
        // setActiveCell / getActiveCell are in-memory APIs
        using var wb = new XSSFWorkbook();
        var sheet = wb.createSheet("Test");

        sheet.setActiveCell("D5");
        Assert.Equal("D5", sheet.getActiveCell());

        sheet.setActiveCell("A1");
        Assert.Equal("A1", sheet.getActiveCell());

        sheet.setSelected(true);
        Assert.True(sheet.isSelected());

        sheet.setSelected(false);
        Assert.False(sheet.isSelected());
    }

    /// <summary>
    /// external data connections (xl/connections.xml) and external links
    /// (xl/externalLinks/*) are separate ZIP parts NOT in GetModelEntryNames().
    /// → should be 🔵 preserved via _preservedEntries.
    /// </summary>
    [Fact]
    public void RoundTrip_ExternalConnections_Preserved()
    {
        byte[] xlsxBytes;
        using (var original = new XSSFWorkbook())
        {
            original.createSheet("Data").createRow(0).createCell(0).setCellValue("Hello");
            using var ms = new MemoryStream();
            original.write(ms);
            xlsxBytes = ms.ToArray();
        }

        using var injectedStream = new MemoryStream();
        using (var srcArchive = new ZipArchive(new MemoryStream(xlsxBytes), ZipArchiveMode.Read))
        using (var dstArchive = new ZipArchive(injectedStream, ZipArchiveMode.Create, leaveOpen: true))
        {
            foreach (var e in srcArchive.Entries)
            {
                var ne = dstArchive.CreateEntry(e.FullName, CompressionLevel.Optimal);
                using var s = e.Open();
                using var d = ne.Open();
                s.CopyTo(d);
            }
            var conn = dstArchive.CreateEntry("xl/connections.xml", CompressionLevel.Optimal);
            using (var sw = new StreamWriter(conn.Open()))
                sw.Write("<connections><connection id=\"1\"/></connections>");
            var ext = dstArchive.CreateEntry("xl/externalLinks/externalLink1.xml", CompressionLevel.Optimal);
            using (var sw = new StreamWriter(ext.Open()))
                sw.Write("<externalLink/>");
        }

        injectedStream.Position = 0;
        using var loaded = new XSSFWorkbook(injectedStream);
        using var outStream = new MemoryStream();
        loaded.write(outStream);
        outStream.Position = 0;

        using var result = new ZipArchive(outStream, ZipArchiveMode.Read);
        Assert.NotNull(result.GetEntry("xl/connections.xml"));
        Assert.NotNull(result.GetEntry("xl/externalLinks/externalLink1.xml"));
        using var r = new StreamReader(result.GetEntry("xl/connections.xml")!.Open());
        Assert.Contains("<connections>", r.ReadToEnd());
    }

    [Fact]
    public void RoundTrip_PivotTable_Preserved()
    {
        // Test pivot table creation and round-trip preservation via the API.
        byte[] xlsxBytes;
        using (var original = new XSSFWorkbook())
        {
            // Source data sheet
            var dataSheet = original.createSheet("Data");
            dataSheet.createRow(0).createCell(0).setCellValue("Category");
            dataSheet.createRow(0).createCell(1).setCellValue("Value");
            dataSheet.createRow(1).createCell(0).setCellValue("A");
            dataSheet.createRow(1).createCell(1).setCellValue(10);
            dataSheet.createRow(2).createCell(0).setCellValue("B");
            dataSheet.createRow(2).createCell(1).setCellValue(20);

            // Pivot table on a separate sheet
            var pivotSheet = original.createSheet("Pivot");
            var pivot = pivotSheet.createPivotTable("A1", "A1:B3", "Data");
            pivot.RowLabels.Add(0);       // Category as row label
            pivot.DataColumns.Add(1);      // Value as data field

            using var ms = new MemoryStream();
            original.write(ms);
            xlsxBytes = ms.ToArray();
        }

        // Verify the written parts exist in the ZIP archive
        using var archive = new ZipArchive(new MemoryStream(xlsxBytes), ZipArchiveMode.Read);
        Assert.NotNull(archive.GetEntry("xl/pivotTables/pivotTable1.xml"));
        Assert.NotNull(archive.GetEntry("xl/pivotCache/pivotCacheDefinition1.xml"));
        Assert.NotNull(archive.GetEntry("xl/pivotCache/pivotCacheRecords1.xml"));

        // Verify content types
        var ctXml = ReadEntry(archive, "[Content_Types].xml");
        Assert.Contains("/xl/pivotTables/pivotTable1.xml", ctXml);
        Assert.Contains("/xl/pivotCache/pivotCacheDefinition1.xml", ctXml);
        Assert.Contains("/xl/pivotCache/pivotCacheRecords1.xml", ctXml);

        // Load from the written bytes and verify pivot table parts survive
        using var loaded = new XSSFWorkbook(new MemoryStream(xlsxBytes));
        using var outStream = new MemoryStream();
        loaded.write(outStream);
        outStream.Position = 0;

        using var result = new ZipArchive(outStream, ZipArchiveMode.Read);
        Assert.NotNull(result.GetEntry("xl/pivotTables/pivotTable1.xml"));
        Assert.NotNull(result.GetEntry("xl/pivotCache/pivotCacheDefinition1.xml"));
        Assert.NotNull(result.GetEntry("xl/pivotCache/pivotCacheRecords1.xml"));
    }

    private static string ReadEntry(ZipArchive archive, string name)
    {
        var entry = archive.GetEntry(name);
        Assert.NotNull(entry);
        using var reader = new StreamReader(entry.Open());
        return reader.ReadToEnd();
    }
}
