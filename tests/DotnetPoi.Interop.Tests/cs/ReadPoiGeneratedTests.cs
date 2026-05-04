using DotnetPoi.SS.UserModel;
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

    [Fact]
    [Trait("Category", "ReadFromPoi")]
    public void Read_FormulaAndBooleanCells_GeneratedByPoi()
    {
        var fixturePath = GetFixturePath("phase7-formulas.xlsx");
        Assert.True(File.Exists(fixturePath), "Run the Java WriteForDotnetTest before this C# read test.");

        using var stream = File.OpenRead(fixturePath);
        using var workbook = new XSSFWorkbook(stream);
        var sheet = workbook.getSheet("Formulas")!;

        // Row 0: base numeric values written by POI
        Assert.Equal(10.0, sheet.getRow(0)!.getCell(0)!.getNumericCellValue());
        Assert.Equal(20.0, sheet.getRow(0)!.getCell(1)!.getNumericCellValue());

        // Row 1: numeric formula SUM → cached value 30
        var formulaNum = sheet.getRow(1)!.getCell(0)!;
        Assert.Equal(CellType.Formula, formulaNum.getCellType());
        Assert.Equal(CellType.Numeric, formulaNum.getCachedFormulaResultType());
        Assert.Equal(30.0, formulaNum.getNumericCellValue());

        // Row 2: string formula → cached "hello world"
        var formulaStr = sheet.getRow(2)!.getCell(0)!;
        Assert.Equal(CellType.Formula, formulaStr.getCellType());
        Assert.Equal(CellType.String, formulaStr.getCachedFormulaResultType());
        Assert.Equal("hello world", formulaStr.getStringCellValue());

        // Row 3: boolean cell (not formula)
        var boolCell = sheet.getRow(3)!.getCell(0)!;
        Assert.Equal(CellType.Boolean, boolCell.getCellType());
        Assert.True(boolCell.getBooleanCellValue());

        // Row 4: error formula → cached #DIV/0!
        var errorCell = sheet.getRow(4)!.getCell(0)!;
        Assert.Equal(CellType.Formula, errorCell.getCellType());
        Assert.Equal(CellType.Error, errorCell.getCachedFormulaResultType());
        Assert.Equal("#DIV/0!", errorCell.getErrorCellString());
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
