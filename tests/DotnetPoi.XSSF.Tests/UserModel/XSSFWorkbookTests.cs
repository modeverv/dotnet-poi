using System.IO.Compression;
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

    private static string ReadEntry(ZipArchive archive, string name)
    {
        var entry = archive.GetEntry(name);
        Assert.NotNull(entry);
        using var reader = new StreamReader(entry.Open());
        return reader.ReadToEnd();
    }
}
