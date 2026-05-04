using DotnetPoi.HSSF.UserModel;

var repoRoot = FindRepositoryRoot(AppContext.BaseDirectory);
var outputDirectory = Path.Combine(repoRoot, "examples", "output");
Directory.CreateDirectory(outputDirectory);

var outputPath = Path.Combine(outputDirectory, "phase4-hssf-xls-example.xls");

using (var workbook = new HSSFWorkbook())
{
    var sheet = workbook.createSheet("Phase4");
    var header = sheet.createRow(0);
    header.createCell(0).setCellValue("Format");
    header.createCell(1).setCellValue("Value");
    header.createCell(2).setCellValue("Supported");

    var xlsRow = sheet.createRow(1);
    xlsRow.createCell(0).setCellValue("HSSF .xls");
    xlsRow.createCell(1).setCellValue(4.0);
    xlsRow.createCell(2).setCellValue(true);

    using var output = File.Create(outputPath);
    workbook.write(output);
}

using (var input = File.OpenRead(outputPath))
using (var workbook = new HSSFWorkbook(input))
{
    var sheet = workbook.getSheet("Phase4") ?? throw new InvalidOperationException("Missing Phase4 sheet.");
    AssertString("HSSF .xls", sheet.getRow(1)?.getCell(0)?.getStringCellValue() ?? "", "format cell");
    AssertDouble(4.0, sheet.getRow(1)?.getCell(1)?.getNumericCellValue() ?? double.NaN, "phase cell");
    AssertBool(true, sheet.getRow(1)?.getCell(2)?.getBooleanCellValue() ?? false, "supported cell");
}

Console.WriteLine("Phase 4 HSSF xls example passed.");
Console.WriteLine($"Wrote {outputPath}");

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

static void AssertBool(bool expected, bool actual, string label)
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
