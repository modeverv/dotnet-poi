using DotnetPoi.SS.UserModel;
using DotnetPoi.XSSF.UserModel;

var repoRoot = FindRepositoryRoot(AppContext.BaseDirectory);
var outputDirectory = Path.Combine(repoRoot, "examples", "output");
Directory.CreateDirectory(outputDirectory);

var outputPath = Path.Combine(outputDirectory, "phase3-interface-example.xlsx");

// Write via IWorkbook interface
WriteViaInterface(outputPath);

// Read back via IWorkbook interface
ReadViaInterface(outputPath);

Console.WriteLine("Phase 3 interface example passed.");
Console.WriteLine($"Output: {outputPath}");

static void WriteViaInterface(string outputPath)
{
    // Typed as IWorkbook — same code will work for any future IWorkbook implementation (e.g. HSSF)
    using IWorkbook workbook = new XSSFWorkbook();

    IFont boldFont = workbook.createFont();
    boldFont.setBold(true);
    boldFont.setFontName("Arial");
    boldFont.setFontHeightInPoints(12);
    boldFont.setColor((short)IndexedColors.DarkBlue);

    ICellStyle headerStyle = workbook.createCellStyle();
    headerStyle.setFont(boldFont);
    headerStyle.setFillForegroundColor((short)IndexedColors.LightTurquoise);
    headerStyle.setFillPattern(FillPatternType.SolidForeground);
    headerStyle.setBorderBottom(BorderStyle.Thin);

    ICellStyle numberStyle = workbook.createCellStyle();
    numberStyle.setDataFormat(workbook.createDataFormat().getFormat("#,##0.00"));

    ISheet sheet = workbook.createSheet("Phase3");

    IRow header = sheet.createRow(0);
    SetHeader(header.createCell(0), "Product", headerStyle);
    SetHeader(header.createCell(1), "Units", headerStyle);
    SetHeader(header.createCell(2), "Price", headerStyle);

    AddDataRow(sheet, 1, "Widget A", 100, 9.99, numberStyle);
    AddDataRow(sheet, 2, "Widget B", 250, 4.50, numberStyle);
    AddDataRow(sheet, 3, "Widget C",  50, 19.99, numberStyle);

    using var stream = File.Create(outputPath);
    workbook.write(stream);
}

static void ReadViaInterface(string outputPath)
{
    using IWorkbook workbook = new XSSFWorkbook(File.OpenRead(outputPath));

    AssertInt(1, workbook.getNumberOfSheets(), "sheet count");

    ISheet sheet = workbook.getSheet("Phase3")
        ?? throw new InvalidOperationException("Sheet 'Phase3' not found.");

    AssertString("Product", sheet.getRow(0)!.getCell(0)!.getStringCellValue(), "A1");
    AssertString("Units",   sheet.getRow(0)!.getCell(1)!.getStringCellValue(), "B1");
    AssertString("Price",   sheet.getRow(0)!.getCell(2)!.getStringCellValue(), "C1");

    AssertString("Widget A", sheet.getRow(1)!.getCell(0)!.getStringCellValue(), "A2");
    AssertDouble(100.0,      sheet.getRow(1)!.getCell(1)!.getNumericCellValue(), "B2");
    AssertDouble(9.99,       sheet.getRow(1)!.getCell(2)!.getNumericCellValue(), "C2");

    AssertInt(3, sheet.getLastRowNum(), "last row num (0-based)");
}

static void SetHeader(ICell cell, string text, ICellStyle style)
{
    cell.setCellValue(text);
    cell.setCellStyle(style);
}

static void AddDataRow(ISheet sheet, int rowIndex, string product, double units, double price, ICellStyle numberStyle)
{
    IRow row = sheet.createRow(rowIndex);
    row.createCell(0).setCellValue(product);

    ICell unitsCell = row.createCell(1);
    unitsCell.setCellValue(units);
    unitsCell.setCellStyle(numberStyle);

    ICell priceCell = row.createCell(2);
    priceCell.setCellValue(price);
    priceCell.setCellStyle(numberStyle);
}

static void AssertString(string expected, string actual, string label)
{
    if (!string.Equals(expected, actual, StringComparison.Ordinal))
        throw new InvalidOperationException($"{label}: expected '{expected}', got '{actual}'.");
}

static void AssertDouble(double expected, double actual, string label)
{
    if (Math.Abs(expected - actual) > 1e-9)
        throw new InvalidOperationException($"{label}: expected {expected}, got {actual}.");
}

static void AssertInt(int expected, int actual, string label)
{
    if (expected != actual)
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
