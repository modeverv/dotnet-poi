using DotnetPoi.SS.UserModel;
using DotnetPoi.XSSF.UserModel;

var repoRoot = FindRepositoryRoot(AppContext.BaseDirectory);
var outputDirectory = Path.Combine(repoRoot, "examples", "output");
Directory.CreateDirectory(outputDirectory);

var outputPath = Path.Combine(outputDirectory, "phase5-formula-evaluator-example.xlsx");

using (var workbook = new XSSFWorkbook())
{
    var sheet = workbook.createSheet("FormulaEvaluator");

    var input = sheet.createRow(0);
    input.createCell(0).setCellValue(10.0);
    input.createCell(1).setCellValue(20.0);
    input.createCell(2).setCellValue(30.0);

    var formulas = sheet.createRow(1);
    formulas.createCell(0).setCellFormula("SUM(A1:C1)");
    formulas.createCell(1).setCellFormula("AVERAGE(A1:C1)");
    formulas.createCell(2).setCellFormula("MIN(A1:C1)");
    formulas.createCell(3).setCellFormula("MAX(A1:C1)");
    formulas.createCell(4).setCellFormula("COUNT(A1:C1)");
    formulas.createCell(5).setCellFormula("CONCATENATE(\"sum=\",SUM(A1:C1))");
    formulas.createCell(6).setCellFormula("\"avg=\"&AVERAGE(A1:C1)");

    var evaluator = workbook.getCreationHelper().createFormulaEvaluator();
    evaluator.evaluateAll();

    using var output = File.Create(outputPath);
    workbook.write(output);
}

using (var stream = File.OpenRead(outputPath))
using (var workbook = new XSSFWorkbook(stream))
{
    var row = workbook.getSheet("FormulaEvaluator")!.getRow(1)!;

    AssertFormula(CellType.Numeric, 60.0, row.getCell(0)!, "SUM");
    AssertFormula(CellType.Numeric, 20.0, row.getCell(1)!, "AVERAGE");
    AssertFormula(CellType.Numeric, 10.0, row.getCell(2)!, "MIN");
    AssertFormula(CellType.Numeric, 30.0, row.getCell(3)!, "MAX");
    AssertFormula(CellType.Numeric, 3.0, row.getCell(4)!, "COUNT");
    AssertFormula(CellType.String, "sum=60", row.getCell(5)!, "CONCATENATE");
    AssertFormula(CellType.String, "avg=20", row.getCell(6)!, "ampersand concat");
}

Console.WriteLine($"Wrote formula evaluator workbook: {outputPath}");

static void AssertFormula(CellType expectedType, object expected, XSSFCell cell, string label)
{
    if (cell.getCellType() != CellType.Formula)
        throw new InvalidOperationException($"{label}: expected formula cell, got {cell.getCellType()}.");
    if (cell.getCachedFormulaResultType() != expectedType)
        throw new InvalidOperationException($"{label}: expected cached {expectedType}, got {cell.getCachedFormulaResultType()}.");

    switch (expectedType)
    {
        case CellType.Numeric:
            var number = cell.getNumericCellValue();
            if (Math.Abs(Convert.ToDouble(expected) - number) > 0.000001)
                throw new InvalidOperationException($"{label}: expected {expected}, got {number}.");
            break;
        case CellType.String:
            var text = cell.getStringCellValue();
            if (!string.Equals((string)expected, text, StringComparison.Ordinal))
                throw new InvalidOperationException($"{label}: expected '{expected}', got '{text}'.");
            break;
        default:
            throw new NotSupportedException($"Unsupported assertion type {expectedType}.");
    }
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
