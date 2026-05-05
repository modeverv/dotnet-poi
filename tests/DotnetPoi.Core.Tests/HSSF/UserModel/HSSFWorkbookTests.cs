using DotnetPoi.HSSF.UserModel;
using DotnetPoi.SS.UserModel;
using Xunit;

namespace DotnetPoi.HSSF.Tests.UserModel;

public sealed class HSSFWorkbookTests
{
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
