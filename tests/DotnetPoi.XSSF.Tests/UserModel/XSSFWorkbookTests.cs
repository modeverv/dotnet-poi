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

    private static string ReadEntry(ZipArchive archive, string name)
    {
        var entry = archive.GetEntry(name);
        Assert.NotNull(entry);
        using var reader = new StreamReader(entry.Open());
        return reader.ReadToEnd();
    }
}
