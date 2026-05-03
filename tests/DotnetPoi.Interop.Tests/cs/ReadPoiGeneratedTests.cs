using DotnetPoi.XSSF.UserModel;
using Xunit;

namespace DotnetPoi.Interop.Tests;

public class ReadPoiGeneratedTests
{
    [Fact]
    [Trait("Category", "ReadFromPoi")]
    public void Read_StringAndNumberWorkbook_GeneratedByPoi()
    {
        var fixturePath = GetFixturePath("phase1-basic.xlsx");
        Assert.True(File.Exists(fixturePath), "Run the Java WriteForDotnetTest before this C# read test.");

        using var stream = File.OpenRead(fixturePath);
        using var workbook = new XSSFWorkbook(stream);

        Assert.Equal(2, workbook.getNumberOfSheets());

        var data = workbook.getSheet("From POI");
        Assert.NotNull(data);
        Assert.Equal("from apache poi", data.getRow(0)!.getCell(0)!.getStringCellValue());
        Assert.Equal(123.25, data.getRow(0)!.getCell(1)!.getNumericCellValue());
        Assert.Equal(0.0, data.getRow(0)!.getCell(2)!.getNumericCellValue());
        Assert.Equal("second row", data.getRow(1)!.getCell(0)!.getStringCellValue());

        var second = workbook.getSheetAt(1);
        Assert.Equal(99.0, second.getRow(2)!.getCell(3)!.getNumericCellValue());
    }

    private static string GetFixturePath(string fileName)
    {
        var directory = AppContext.BaseDirectory;
        while (directory is not null)
        {
            var fixtures = Path.Combine(directory, "tests", "DotnetPoi.Interop.Tests", "fixtures");
            if (Directory.Exists(fixtures))
            {
                return Path.Combine(fixtures, "from-poi", fileName);
            }

            directory = Directory.GetParent(directory)?.FullName;
        }

        throw new DirectoryNotFoundException("Could not locate tests/DotnetPoi.Interop.Tests/fixtures.");
    }
}
