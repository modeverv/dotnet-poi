using DotnetPoi.XSSF.UserModel;
using Xunit;

namespace DotnetPoi.Interop.Tests;

public class WriteForPoiTests
{
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
