using DotnetPoi.SS.UserModel;
using DotnetPoi.XSSF.UserModel;

// ------------------------------------------------------------------
// Phase 7 — Cell type coverage example
//
// Demonstrates reading all cell types from an xlsx that was generated
// by Apache POI (Java): numeric, string, boolean, and formula cells
// (with their cached result: numeric, string, boolean, or error).
// ------------------------------------------------------------------

var repoRoot = FindRepositoryRoot(AppContext.BaseDirectory);
var fixturePath = Path.Combine(repoRoot,
    "tests", "DotnetPoi.Interop.Tests", "fixtures", "from-poi", "phase7-formulas.xlsx");

if (!File.Exists(fixturePath))
{
    Console.WriteLine("Fixture not found. Generate it first:");
    Console.WriteLine("  mvn test -f tests/DotnetPoi.Interop.Tests/java/pom.xml -Dtest=WriteForDotnetTest");
    return;
}

using var stream = File.OpenRead(fixturePath);
using var workbook = new XSSFWorkbook(stream);
var sheet = workbook.getSheet("Formulas")!;

Console.WriteLine("Reading POI-generated xlsx with formula/boolean/error cells:");
Console.WriteLine();

for (int r = 0; r <= 4; r++)
{
    var row = sheet.getRow(r);
    if (row is null) continue;
    var cell = row.getCell(0);
    if (cell is null) continue;

    var display = cell.getCellType() switch
    {
        CellType.Numeric => $"NUMERIC        = {cell.getNumericCellValue()}",
        CellType.String  => $"STRING         = \"{cell.getStringCellValue()}\"",
        CellType.Boolean => $"BOOLEAN        = {cell.getBooleanCellValue()}",
        CellType.Error   => $"ERROR          = {cell.getErrorCellString()}",
        CellType.Formula => cell.getCachedFormulaResultType() switch
        {
            CellType.Numeric => $"FORMULA→NUM    = {cell.getNumericCellValue()}",
            CellType.String  => $"FORMULA→STR    = \"{cell.getStringCellValue()}\"",
            CellType.Boolean => $"FORMULA→BOOL   = {cell.getBooleanCellValue()}",
            CellType.Error   => $"FORMULA→ERROR  = {cell.getErrorCellString()}",
            var t            => $"FORMULA→{t}"
        },
        var t => $"{t}"
    };

    Console.WriteLine($"  row {r}: {display}");
}

Console.WriteLine();
Console.WriteLine("Phase 7 cell-types example passed.");

// Demonstrate write round-trip with Boolean cells
var outputPath = Path.Combine(repoRoot, "examples", "output", "phase7-cell-types-example.xlsx");
Directory.CreateDirectory(Path.GetDirectoryName(outputPath)!);

using var wb = new XSSFWorkbook();
var ws = wb.createSheet("Types");
var wr = ws.createRow(0);
wr.createCell(0).setCellValue(true);
wr.createCell(1).setCellValue(false);
wr.createCell(2).setCellValue(99.9);
wr.createCell(3).setCellValue("hello");

using (var outStream = File.Create(outputPath))
{
    wb.write(outStream);
}

// Read back
using var readStream = File.OpenRead(outputPath);
using var wb2 = new XSSFWorkbook(readStream);
var ws2 = wb2.getSheetAt(0);
Assert(CellType.Boolean, ws2.getRow(0)!.getCell(0)!.getCellType(), "A1 type");
Assert(true,  ws2.getRow(0)!.getCell(0)!.getBooleanCellValue(), "A1 bool");
Assert(false, ws2.getRow(0)!.getCell(1)!.getBooleanCellValue(), "B1 bool");
Assert(99.9,  ws2.getRow(0)!.getCell(2)!.getNumericCellValue(), "C1 numeric");
Assert("hello", ws2.getRow(0)!.getCell(3)!.getStringCellValue(), "D1 string");

Console.WriteLine($"Write round-trip verified: {outputPath}");

static void Assert<T>(T expected, T actual, string label)
{
    if (!EqualityComparer<T>.Default.Equals(expected, actual))
        throw new InvalidOperationException($"{label}: expected {expected}, got {actual}.");
}

static string FindRepositoryRoot(string startDirectory)
{
    var directory = new DirectoryInfo(startDirectory);
    while (directory is not null)
    {
        if (File.Exists(Path.Combine(directory.FullName, "DotnetPOI.sln")))
            return directory.FullName;
        directory = directory.Parent;
    }
    throw new DirectoryNotFoundException("Could not locate the dotnet-poi repository root.");
}
