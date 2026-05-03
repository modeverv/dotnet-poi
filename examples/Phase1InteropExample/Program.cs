using DotnetPoi.XSSF.UserModel;

var repoRoot = FindRepositoryRoot(AppContext.BaseDirectory);
var outputDirectory = Path.Combine(repoRoot, "examples", "output");
Directory.CreateDirectory(outputDirectory);

var dotnetPoiOutputPath = Path.Combine(outputDirectory, "phase1-dotnet-poi-write.xlsx");
WriteDotnetPoiWorkbook(dotnetPoiOutputPath);
ReadAndAssertDotnetPoiWorkbook(dotnetPoiOutputPath);

var poiFixturePath = Path.Combine(repoRoot, "tests", "DotnetPoi.Interop.Tests", "fixtures", "from-poi", "phase1-basic.xlsx");
if (!File.Exists(poiFixturePath))
{
    Console.WriteLine("POI fixture is missing.");
    Console.WriteLine("Generate it with:");
    Console.WriteLine("mvn test -f tests/DotnetPoi.Interop.Tests/java/pom.xml -Dtest=WriteForDotnetTest");
    return;
}

ReadAndAssertPoiWorkbook(poiFixturePath);

Console.WriteLine("Phase 1 interop example passed.");
Console.WriteLine($"dotnet-poi wrote: {dotnetPoiOutputPath}");
Console.WriteLine($"dotnet-poi read POI fixture: {poiFixturePath}");

static void WriteDotnetPoiWorkbook(string outputPath)
{
    using var workbook = new XSSFWorkbook();
    var sheet = workbook.createSheet("Dotnet");
    var row = sheet.createRow(0);
    row.createCell(0).setCellValue("from dotnet-poi example");
    row.createCell(1).setCellValue(456.75);
    row.createCell(2).setCellValue(0.0);

    var secondRow = sheet.createRow(2);
    secondRow.createCell(3).setCellValue("sparse cell");

    using var stream = File.Create(outputPath);
    workbook.write(stream);
}

static void ReadAndAssertDotnetPoiWorkbook(string path)
{
    using var stream = File.OpenRead(path);
    using var workbook = new XSSFWorkbook(stream);

    AssertInt(1, workbook.getNumberOfSheets(), "dotnet workbook sheet count");
    var sheet = workbook.getSheet("Dotnet") ?? throw new InvalidOperationException("Missing Dotnet sheet.");
    AssertString("from dotnet-poi example", sheet.getRow(0)!.getCell(0)!.getStringCellValue(), "dotnet A1");
    AssertDouble(456.75, sheet.getRow(0)!.getCell(1)!.getNumericCellValue(), "dotnet B1");
    AssertDouble(0.0, sheet.getRow(0)!.getCell(2)!.getNumericCellValue(), "dotnet C1");
    AssertString("sparse cell", sheet.getRow(2)!.getCell(3)!.getStringCellValue(), "dotnet D3");
}

static void ReadAndAssertPoiWorkbook(string path)
{
    using var stream = File.OpenRead(path);
    using var workbook = new XSSFWorkbook(stream);

    AssertInt(2, workbook.getNumberOfSheets(), "POI workbook sheet count");
    var data = workbook.getSheet("From POI") ?? throw new InvalidOperationException("Missing From POI sheet.");
    AssertString("from apache poi", data.getRow(0)!.getCell(0)!.getStringCellValue(), "POI A1");
    AssertDouble(123.25, data.getRow(0)!.getCell(1)!.getNumericCellValue(), "POI B1");
    AssertDouble(0.0, data.getRow(0)!.getCell(2)!.getNumericCellValue(), "POI C1");
    AssertString("second row", data.getRow(1)!.getCell(0)!.getStringCellValue(), "POI A2");

    var second = workbook.getSheetAt(1);
    AssertDouble(99.0, second.getRow(2)!.getCell(3)!.getNumericCellValue(), "POI D3");
}

static void AssertString(string expected, string actual, string label)
{
    if (!string.Equals(expected, actual, StringComparison.Ordinal))
    {
        throw new InvalidOperationException($"{label}: expected '{expected}', got '{actual}'.");
    }
}

static void AssertDouble(double expected, double actual, string label)
{
    if (Math.Abs(expected - actual) > 0.0000001)
    {
        throw new InvalidOperationException($"{label}: expected {expected}, got {actual}.");
    }
}

static void AssertInt(int expected, int actual, string label)
{
    if (expected != actual)
    {
        throw new InvalidOperationException($"{label}: expected {expected}, got {actual}.");
    }
}

static string FindRepositoryRoot(string startDirectory)
{
    var directory = new DirectoryInfo(startDirectory);
    while (directory is not null)
    {
        if (File.Exists(Path.Combine(directory.FullName, "DotnetPOI.sln")))
        {
            return directory.FullName;
        }

        directory = directory.Parent;
    }

    throw new DirectoryNotFoundException("Could not locate the dotnet-poi repository root.");
}
