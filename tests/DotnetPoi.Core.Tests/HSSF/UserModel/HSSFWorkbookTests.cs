using DotnetPoi.HSSF.UserModel;
using DotnetPoi.POIFS.Crypt;
using DotnetPoi.SS.UserModel;
using Xunit;

namespace DotnetPoi.HSSF.Tests.UserModel;

public sealed class HSSFWorkbookTests
{
    public static IEnumerable<object[]> Phase12RepresentativePoiFixtures()
    {
        yield return new object[] { "empty.xls", 3 };
        yield return new object[] { "Simple.xls", 3 };
        yield return new object[] { "SimpleMultiCell.xls", 3 };
        yield return new object[] { "SampleSS.xls", 3 };
        yield return new object[] { "WORKBOOK_in_capitals.xls", 1 };
        yield return new object[] { "BOOK_in_capitals.xls", 1 };
        yield return new object[] { "chinese-provinces.xls", 1 };
        yield return new object[] { "DateFormats.xls", 3 };
        yield return new object[] { "SimpleWithStyling.xls", 3 };
        yield return new object[] { "WithExtendedStyles.xls", 3 };
        yield return new object[] { "55341_CellStyleBorder.xls", 3 };
        yield return new object[] { "SimpleWithPrintArea.xls", 3 };
        yield return new object[] { "RepeatingRowsCols.xls", 4 };
        yield return new object[] { "SimpleWithFormula.xls", 3 };
        yield return new object[] { "ex47747-sharedFormula.xls", 1 };
        yield return new object[] { "WithHyperlink.xls", 3 };
        yield return new object[] { "comments.xls", 3 };
        yield return new object[] { "drawings.xls", 8 };
        yield return new object[] { "SimpleWithImages.xls", 3 };
        yield return new object[] { "SimpleMacro.xls", 3 };
    }

    [Fact]
    public void Write_StringNumberBooleanCells_RoundTrips()
    {
        using var workbook = new HSSFWorkbook();
        var sheet = workbook.createSheet("Data");
        var row = sheet.createRow(0);
        row.createCell(0).setCellValue("alpha");
        row.createCell(1).setCellValue(42.5);
        row.createCell(2).setCellValue(true);

        using var stream = new MemoryStream();
        workbook.write(stream);
        stream.Position = 0;

        using var read = new HSSFWorkbook(stream);

        Assert.Equal(1, read.getNumberOfSheets());
        var readRow = read.getSheetAt(0).getRow(0);
        Assert.NotNull(readRow);
        Assert.Equal("alpha", readRow!.getCell(0)!.getStringCellValue());
        Assert.Equal(42.5, readRow.getCell(1)!.getNumericCellValue());
        Assert.True(readRow.getCell(2)!.getBooleanCellValue());
    }

    [Fact]
    public void Read_PoiSampleXls_LoadsWorkbookStreamAndBasicCells()
    {
        var sample = FindRepoRoot().Combine("poi/test-data/spreadsheet/SimpleMultiCell.xls");
        using var input = sample.OpenRead();

        using var workbook = new HSSFWorkbook(input);

        Assert.True(workbook.getNumberOfSheets() > 0);
        var sheet = workbook.getSheetAt(0);
        Assert.True(sheet.getLastRowNum() >= 0);
        Assert.Contains(Enumerable.Range(0, sheet.getLastRowNum() + 1), rowIndex =>
        {
            var row = sheet.getRow(rowIndex);
            return row is not null && Enumerable.Range(0, Math.Max(row.getLastCellNum(), (short)0))
                .Select(row.getCell)
                .Any(cell => cell is not null && cell.getCellType() is CellType.String or CellType.Numeric or CellType.Boolean);
        });
    }

    [Theory]
    [MemberData(nameof(Phase12RepresentativePoiFixtures))]
    public void Read_Phase12RepresentativePoiFixtures_LoadsWorkbookStream(string fileName, int expectedSheetCount)
    {
        var sample = FindRepoRoot().Combine(Path.Combine("poi/test-data/spreadsheet", fileName));
        using var input = sample.OpenRead();

        using var workbook = new HSSFWorkbook(input);

        Assert.Equal(expectedSheetCount, workbook.getNumberOfSheets());
    }

    [Fact]
    public void Read_BookUppercasePoiFixture_LoadsWorkbookStream()
    {
        var sample = FindRepoRoot().Combine("poi/test-data/spreadsheet/BOOK_in_capitals.xls");
        using var input = sample.OpenRead();

        using var workbook = new HSSFWorkbook(input);

        Assert.Equal(1, workbook.getNumberOfSheets());
    }

    [Fact]
    public void Write_LoadedPoiFixture_PreservesNonWorkbookOleStreams()
    {
        var sample = FindRepoRoot().Combine("poi/test-data/spreadsheet/empty.xls");
        Dictionary<string, byte[]> originalStreams;
        using (var originalInput = sample.OpenRead())
        {
            originalStreams = CompoundFile.ReadStreams(originalInput);
        }

        using var workbookInput = sample.OpenRead();
        using var workbook = new HSSFWorkbook(workbookInput);

        using var output = new MemoryStream();
        workbook.write(output);
        output.Position = 0;
        var writtenStreams = CompoundFile.ReadStreams(output);

        foreach (var streamName in originalStreams.Keys.Where(name => name != "Workbook"))
        {
            Assert.True(writtenStreams.ContainsKey(streamName), $"Missing preserved stream '{streamName}'.");
            Assert.Equal(originalStreams[streamName], writtenStreams[streamName]);
        }
    }

    [Fact]
    public void Write_BookUppercasePoiFixture_PreservesWorkbookStreamAlias()
    {
        var sample = FindRepoRoot().Combine("poi/test-data/spreadsheet/BOOK_in_capitals.xls");
        using var input = sample.OpenRead();
        using var workbook = new HSSFWorkbook(input);

        using var output = new MemoryStream();
        workbook.write(output);
        output.Position = 0;
        var writtenStreams = CompoundFile.ReadStreams(output);

        Assert.True(writtenStreams.ContainsKey("BOOK"));
        Assert.False(writtenStreams.ContainsKey("Workbook"));
        Assert.False(writtenStreams.ContainsKey("Book"));
        Assert.False(writtenStreams.ContainsKey("WORKBOOK"));
    }

    private static DirectoryInfo FindRepoRoot()
    {
        var dir = new DirectoryInfo(AppContext.BaseDirectory);
        while (dir is not null && !File.Exists(Path.Combine(dir.FullName, "DotnetPOI.sln")))
        {
            dir = dir.Parent;
        }

        return dir ?? throw new InvalidOperationException("Could not locate repository root.");
    }
}

internal static class DirectoryInfoExtensions
{
    public static FileInfo Combine(this DirectoryInfo directory, string relativePath) =>
        new(Path.Combine(directory.FullName, relativePath));
}
