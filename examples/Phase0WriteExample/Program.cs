using DotnetPoi.XSSF.UserModel;

var repoRoot = FindRepositoryRoot(AppContext.BaseDirectory);
var outputDirectory = Path.Combine(repoRoot, "examples", "output");
Directory.CreateDirectory(outputDirectory);

var outputPath = Path.Combine(outputDirectory, "phase0-write-example.xlsx");

using var workbook = new XSSFWorkbook();
var sheet = workbook.createSheet("Phase0");

var header = sheet.createRow(0);
header.createCell(0).setCellValue("Item");
header.createCell(1).setCellValue("Value");
header.createCell(2).setCellValue("Note");

var stringRow = sheet.createRow(1);
stringRow.createCell(0).setCellValue("String cell");
stringRow.createCell(1).setCellValue("Hello from dotnet-poi");
stringRow.createCell(2).setCellValue("Written through sharedStrings.xml");

var numberRow = sheet.createRow(2);
numberRow.createCell(0).setCellValue("Number cell");
numberRow.createCell(1).setCellValue(123.25);
numberRow.createCell(2).setCellValue("Written as a numeric <v> value");

var zeroRow = sheet.createRow(3);
zeroRow.createCell(0).setCellValue("Zero cell");
zeroRow.createCell(1).setCellValue(0.0);
zeroRow.createCell(2).setCellValue("Keeps explicit zero output");

using (var stream = File.Create(outputPath))
{
    workbook.write(stream);
}

Console.WriteLine($"Wrote {outputPath}");

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
