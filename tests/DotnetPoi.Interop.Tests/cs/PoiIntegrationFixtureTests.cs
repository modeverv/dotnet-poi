using DotnetPoi.SS.UserModel;
using DotnetPoi.XSSF.UserModel;
using System.IO.Compression;
using Xunit;

namespace DotnetPoi.Interop.Tests;

/// <summary>
/// Semantic XSSF tests using the POI integration fixture packages (item 10).
///
/// These tests read xlsx/xlsm files produced by Apache POI and verify that
/// DotnetPoi parses cell values, shared strings, styles, relationships, and
/// macro-enabled workbook content correctly.
///
/// Fixture source: tests/DotnetPoi.Interop.Tests/fixtures/poi-integration/_workbooks/
/// Fixture generator: tests/DotnetPoi.Interop.Tests/java/src/test/java/.../PoiIntegrationFixtureGeneratorTest.java
///
/// Item 11 reminder: if these tests expose a lexical mismatch in PoiXmlWriter,
/// add a focused PoiXmlWriter test for that specific behavior before fixing it.
/// Item 12 reminder: do not resurrect fixture-specific XSSFWorkbook XML payloads;
/// keep writers model-driven and POI-source-backed.
/// </summary>
public class PoiIntegrationFixtureTests
{
    // ── shared-strings-basic ──────────────────────────────────────────────

    [Fact]
    public void Read_SharedStringsBasic_ThreeSheetsWithCorrectNames()
    {
        using var workbook = OpenFixture("poi-integration-shared-strings-basic.xlsx");

        Assert.Equal(3, workbook.getNumberOfSheets());
        Assert.NotNull(workbook.getSheet("Sheet1"));
        Assert.NotNull(workbook.getSheet("rich test"));
        Assert.NotNull(workbook.getSheet("Sheet3"));
    }

    [Fact]
    public void Read_SharedStringsBasic_Sheet1StringAndNumericCellValues()
    {
        using var workbook = OpenFixture("poi-integration-shared-strings-basic.xlsx");
        var sheet = workbook.getSheet("Sheet1")!;

        // getRow uses 0-based indices: getRow(0) = spreadsheet row 1
        // A1 = "Lorem" (shared string index 0)
        Assert.Equal("Lorem", sheet.getRow(0)!.getCell(0)!.getStringCellValue());
        // B1 = 111 (numeric)
        Assert.Equal(111.0, sheet.getRow(0)!.getCell(1)!.getNumericCellValue());

        // A2 = "ipsum"
        Assert.Equal("ipsum", sheet.getRow(1)!.getCell(0)!.getStringCellValue());
        Assert.Equal(222.0, sheet.getRow(1)!.getCell(1)!.getNumericCellValue());
    }

    // ── shared-strings-escaping ───────────────────────────────────────────

    [Fact]
    public void Read_SharedStringsEscaping_XmlEntitiesDecodedToLiteralChars()
    {
        using var workbook = OpenFixture("poi-integration-shared-strings-escaping.xlsx");
        var sheet = workbook.getSheetAt(0);

        // getRow(0) = spreadsheet row 1.
        // Shared string index 0 contains literal "<" stored as &lt; in XML.
        var cellValue = sheet.getRow(0)!.getCell(0)!.getStringCellValue();
        Assert.Contains("<", cellValue);
        Assert.DoesNotContain("&lt;", cellValue);
    }

    // ── styles-formatting ────────────────────────────────────────────────

    [Fact]
    public void Read_StylesFormatting_ThreeSheetsExist()
    {
        using var workbook = OpenFixture("poi-integration-styles-formatting.xlsx");

        Assert.Equal(3, workbook.getNumberOfSheets());
        Assert.NotNull(workbook.getSheet("Sheet1"));
        Assert.NotNull(workbook.getSheet("Sheet2"));
        Assert.NotNull(workbook.getSheet("Sheet3"));
    }

    [Fact]
    public void Read_StylesFormatting_Sheet1HasStringCellsInColumnA()
    {
        using var workbook = OpenFixture("poi-integration-styles-formatting.xlsx");
        var sheet = workbook.getSheet("Sheet1")!;

        // getRow(0) = spreadsheet row 1.
        // A1 = shared string index 6 = "Dates, all 24th November 2006"
        var cell = sheet.getRow(0)?.getCell(0);
        Assert.NotNull(cell);
        Assert.Equal(CellType.String, cell!.getCellType());
        Assert.Equal("Dates, all 24th November 2006", cell.getStringCellValue());
    }

    // ── comments-write-read ───────────────────────────────────────────────

    [Fact]
    public void Read_CommentsWriteRead_ThreeSheetsWithExpectedNames()
    {
        using var workbook = OpenFixture("poi-integration-comments-write-read.xlsx");

        // The fixture has 3 sheets; "AllANumbers" and "AllBStrings" are defined names, not sheets.
        Assert.Equal(3, workbook.getNumberOfSheets());
        Assert.NotNull(workbook.getSheet("Sheet1"));
        Assert.NotNull(workbook.getSheet("Sheet2"));
        Assert.NotNull(workbook.getSheet("Sheet3"));
    }

    [Fact]
    public void Read_CommentsWriteRead_Sheet1ColumnACellsAreStrings()
    {
        using var workbook = OpenFixture("poi-integration-comments-write-read.xlsx");
        var sheet = workbook.getSheet("Sheet1")!;

        // getRow(0) = spreadsheet row 1.
        // A1 = shared string index 0 = "A1"
        Assert.Equal("A1", sheet.getRow(0)!.getCell(0)!.getStringCellValue());
        // B1 = shared string index 1 = "B1"
        Assert.Equal("B1", sheet.getRow(0)!.getCell(1)!.getStringCellValue());
    }

    [Fact]
    public void Read_CommentsWriteRead_Sheet1NumericCellsInColumnA()
    {
        using var workbook = OpenFixture("poi-integration-comments-write-read.xlsx");
        var sheet = workbook.getSheet("Sheet1")!;

        // getRow(1) = spreadsheet row 2. A2 = 22.3 (numeric, no t="s" in XML).
        Assert.Equal(22.3, sheet.getRow(1)!.getCell(0)!.getNumericCellValue());
        // getRow(2) = spreadsheet row 3. A3 = 24.5
        Assert.Equal(24.5, sheet.getRow(2)!.getCell(0)!.getNumericCellValue());
    }

    // ── xlsm-vba-preserve ────────────────────────────────────────────────

    [Fact]
    public void Read_XlsmVbaPreserve_HasMacrosIsTrue()
    {
        using var workbook = OpenFixture("poi-integration-xlsm-vba-preserve.xlsm");
        Assert.True(workbook.HasMacros);
    }

    [Fact]
    public void Read_XlsmVbaPreserve_ThreeSheetsWithExpectedNames()
    {
        using var workbook = OpenFixture("poi-integration-xlsm-vba-preserve.xlsm");

        Assert.Equal(3, workbook.getNumberOfSheets());
        Assert.NotNull(workbook.getSheet("SheetA"));
        Assert.NotNull(workbook.getSheet("SheetB"));
        Assert.NotNull(workbook.getSheet("SheetC"));
    }

    [Fact]
    public void RoundTrip_XlsmVbaPreserve_VbaBytesPreservedByteForByte()
    {
        // Read the original fixture and capture the VBA bytes.
        byte[] originalVbaBytes;
        using (var fixture = OpenFixtureStream("poi-integration-xlsm-vba-preserve.xlsm"))
        using (var originalZip = new ZipArchive(fixture, ZipArchiveMode.Read))
        {
            var entry = originalZip.GetEntry("xl/vbaProject.bin")
                ?? throw new InvalidDataException("xl/vbaProject.bin missing from fixture");
            using var s = entry.Open();
            using var ms = new MemoryStream();
            s.CopyTo(ms);
            originalVbaBytes = ms.ToArray();
        }

        // Round-trip: read → write → read vbaProject.bin.
        byte[] roundTrippedVbaBytes;
        using (var original = OpenFixture("poi-integration-xlsm-vba-preserve.xlsm"))
        using (var roundTripped = new MemoryStream())
        {
            original.write(roundTripped);
            roundTripped.Position = 0;
            using var rtZip = new ZipArchive(roundTripped, ZipArchiveMode.Read);
            var entry = rtZip.GetEntry("xl/vbaProject.bin")
                ?? throw new InvalidDataException("xl/vbaProject.bin missing from round-tripped xlsm");
            using var s = entry.Open();
            using var ms = new MemoryStream();
            s.CopyTo(ms);
            roundTrippedVbaBytes = ms.ToArray();
        }

        Assert.Equal(originalVbaBytes, roundTrippedVbaBytes);
    }

    // ── Helper ────────────────────────────────────────────────────────────

    private static XSSFWorkbook OpenFixture(string fileName)
    {
        var stream = OpenFixtureStream(fileName);
        try
        {
            return new XSSFWorkbook(stream);
        }
        finally
        {
            stream.Dispose();
        }
    }

    private static Stream OpenFixtureStream(string fileName)
    {
        var path = GetPoiIntegrationFixturePath(fileName);
        Assert.True(File.Exists(path),
            $"POI integration fixture not found at {path}. " +
            "Run: mvn test -f tests/DotnetPoi.Interop.Tests/java/pom.xml -Dtest=PoiIntegrationFixtureGeneratorTest");
        return File.OpenRead(path);
    }

    private static string GetPoiIntegrationFixturePath(string fileName)
    {
        var directory = AppContext.BaseDirectory;
        while (directory is not null)
        {
            var workbooksDir = Path.Combine(
                directory,
                "tests",
                "DotnetPoi.Interop.Tests",
                "fixtures",
                "poi-integration",
                "_workbooks");
            if (Directory.Exists(workbooksDir))
            {
                return Path.Combine(workbooksDir, fileName);
            }

            directory = Directory.GetParent(directory)?.FullName;
        }

        throw new DirectoryNotFoundException(
            "Could not locate tests/DotnetPoi.Interop.Tests/fixtures/poi-integration/_workbooks.");
    }
}
