using System.IO.Compression;
using DotnetPoi.SS.UserModel;
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

    private static string ReadEntry(ZipArchive archive, string name)
    {
        var entry = archive.GetEntry(name);
        Assert.NotNull(entry);
        using var reader = new StreamReader(entry.Open());
        return reader.ReadToEnd();
    }
}
