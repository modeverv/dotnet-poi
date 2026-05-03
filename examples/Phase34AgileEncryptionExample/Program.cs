using DotnetPoi.POIFS.Crypt;
using DotnetPoi.XSSF.UserModel;

const string password = "f";

var repoRoot = FindRepositoryRoot(AppContext.BaseDirectory);
var outputDirectory = Path.Combine(repoRoot, "examples", "output");
Directory.CreateDirectory(outputDirectory);

var outputPath = Path.Combine(outputDirectory, "phase3_4-agile-encrypted-example.xlsx");

using (var workbook = new XSSFWorkbook())
{
    var sheet = workbook.createSheet("Phase3.4");
    var header = sheet.createRow(0);
    header.createCell(0).setCellValue("Feature");
    header.createCell(1).setCellValue("Status");

    var row = sheet.createRow(1);
    row.createCell(0).setCellValue("OOXML Agile encryption");
    row.createCell(1).setCellValue("Encrypted by dotnet-poi");

    var numericRow = sheet.createRow(2);
    numericRow.createCell(0).setCellValue("Phase");
    numericRow.createCell(1).setCellValue(3.4);

    using var output = File.Create(outputPath);
    workbook.writeEncrypted(output, password);
}

using (var encrypted = File.OpenRead(outputPath))
{
    var info = new EncryptionInfo(encrypted);
    if (!info.Decryptor.verifyPassword(password))
    {
        throw new InvalidOperationException("Password verification failed.");
    }

    using var decrypted = new MemoryStream(info.Decryptor.getData());
    using var roundTrip = new XSSFWorkbook(decrypted);
    var sheet = roundTrip.getSheet("Phase3.4") ?? throw new InvalidOperationException("Missing Phase3.4 sheet.");
    AssertString("OOXML Agile encryption", sheet.getRow(1)?.getCell(0)?.getStringCellValue() ?? "", "feature cell");
    AssertString("Encrypted by dotnet-poi", sheet.getRow(1)?.getCell(1)?.getStringCellValue() ?? "", "status cell");
    AssertDouble(3.4, sheet.getRow(2)?.getCell(1)?.getNumericCellValue() ?? double.NaN, "phase cell");
}

Console.WriteLine($"Wrote encrypted workbook: {outputPath}");
Console.WriteLine($"Password: {password}");

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

static void AssertString(string expected, string actual, string label)
{
    if (!string.Equals(expected, actual, StringComparison.Ordinal))
    {
        throw new InvalidOperationException($"{label}: expected '{expected}', got '{actual}'.");
    }
}

static void AssertDouble(double expected, double actual, string label)
{
    if (Math.Abs(expected - actual) > 0.000001)
    {
        throw new InvalidOperationException($"{label}: expected {expected}, got {actual}.");
    }
}
