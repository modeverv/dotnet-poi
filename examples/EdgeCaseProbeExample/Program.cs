using DotnetPoi.SS.UserModel;
using DotnetPoi.XSLF.UserModel;
using DotnetPoi.XSSF.UserModel;
using DotnetPoi.XWPF.UserModel;

var repoRoot = FindRepositoryRoot(AppContext.BaseDirectory);
var outputDirectory = Path.Combine(repoRoot, "examples", "output");
Directory.CreateDirectory(outputDirectory);

var results = new List<ProbeResult>();

Probe("xlsx empty workbook auto-creates default sheet", () =>
{
    var path = Path.Combine(outputDirectory, "edge-empty-workbook.xlsx");
    using (var workbook = new XSSFWorkbook())
    using (var output = File.Create(path))
    {
        workbook.write(output);
    }

    using var input = File.OpenRead(path);
    using var read = new XSSFWorkbook(input);
    AssertEqual(1, read.getNumberOfSheets(), "sheet count");
    AssertNotNull(read.getSheetAt(0), "default sheet");
});

Probe("xlsx sparse rows, sparse cells, XML-sensitive and unicode strings", () =>
{
    var path = Path.Combine(outputDirectory, "edge-sparse-and-strings.xlsx");
    const string xmlSensitive = "<tag attr=\"&\">line1\nline2\t'end'</tag>";
    const string unicode = "\u65e5\u672c\u8a9e\u3068\ud83d\ude80";

    using (var workbook = new XSSFWorkbook())
    {
        var sheet = workbook.createSheet("Sparse");
        sheet.createRow(0).createCell(0).setCellValue("");
        sheet.createRow(999).createCell(51).setCellValue(xmlSensitive);
        sheet.createRow(1000).createCell(52).setCellValue(unicode);
        sheet.createRow(1001).createCell(0).setCellValue(-1234567890.125);

        using var output = File.Create(path);
        workbook.write(output);
    }

    using var input = File.OpenRead(path);
    using var read = new XSSFWorkbook(input);
    var sheet2 = read.getSheet("Sparse")!;
    AssertEqual(1001, sheet2.getLastRowNum(), "last row");
    AssertEqual(xmlSensitive, sheet2.getRow(999)!.getCell(51)!.getStringCellValue(), "XML-sensitive string");
    AssertEqual(unicode, sheet2.getRow(1000)!.getCell(52)!.getStringCellValue(), "unicode string");
    AssertNear(-1234567890.125, sheet2.getRow(1001)!.getCell(0)!.getNumericCellValue(), "negative decimal");
});

Probe("xlsx formula edge results: divide by zero, circular reference, missing cells", () =>
{
    var path = Path.Combine(outputDirectory, "edge-formulas.xlsx");
    using (var workbook = new XSSFWorkbook())
    {
        var sheet = workbook.createSheet("FormulaEdges");
        var row = sheet.createRow(0);
        row.createCell(0).setCellValue(10.0);
        row.createCell(1).setCellFormula("A1/0");
        row.createCell(2).setCellFormula("C1+1");
        row.createCell(3).setCellFormula("SUM(A1:A2,Z99)");
        row.createCell(4).setCellFormula("Z99+2");

        var evaluator = workbook.getCreationHelper().createFormulaEvaluator();
        evaluator.evaluateAll();

        AssertEqual(CellType.Error, row.getCell(1)!.getCachedFormulaResultType(), "divide by zero cached type");
        AssertEqual("#DIV/0!", row.getCell(1)!.getErrorCellString(), "divide by zero cached value");
        AssertEqual(CellType.Error, row.getCell(2)!.getCachedFormulaResultType(), "circular cached type");
        AssertEqual(CellType.Numeric, row.getCell(3)!.getCachedFormulaResultType(), "range with missing cells cached type");
        AssertNear(10.0, row.getCell(3)!.getNumericCellValue(), "range with missing cells value");
        AssertNear(2.0, row.getCell(4)!.getNumericCellValue(), "missing ref as zero");

        using var output = File.Create(path);
        workbook.write(output);
    }

    using var input = File.OpenRead(path);
    using var read = new XSSFWorkbook(input);
    var row2 = read.getSheet("FormulaEdges")!.getRow(0)!;
    AssertEqual("#DIV/0!", row2.getCell(1)!.getErrorCellString(), "read divide by zero");
    AssertEqual(CellType.Error, row2.getCell(2)!.getCachedFormulaResultType(), "read circular cached type");
    AssertNear(10.0, row2.getCell(3)!.getNumericCellValue(), "read range sum");
});

Probe("xlsx encrypted sparse workbook round-trips through decryptor", () =>
{
    var path = Path.Combine(outputDirectory, "edge-encrypted-sparse.xlsx");
    using (var workbook = new XSSFWorkbook())
    {
        var sheet = workbook.createSheet("EncryptedSparse");
        sheet.createRow(250).createCell(8).setCellValue("secret edge");
        sheet.createRow(251).createCell(8).setCellValue(0.0);

        using var output = File.Create(path);
        workbook.writeEncrypted(output, "edge-pass");
    }

    using var encrypted = File.OpenRead(path);
    var info = new DotnetPoi.POIFS.Crypt.EncryptionInfo(encrypted);
    if (!info.Decryptor.verifyPassword("edge-pass"))
        throw new InvalidOperationException("Password verification failed.");

    using var decrypted = new MemoryStream(info.Decryptor.getData());
    using var read = new XSSFWorkbook(decrypted);
    var sheet2 = read.getSheet("EncryptedSparse")!;
    AssertEqual("secret edge", sheet2.getRow(250)!.getCell(8)!.getStringCellValue(), "encrypted string");
    AssertNear(0.0, sheet2.getRow(251)!.getCell(8)!.getNumericCellValue(), "encrypted zero");
});

Probe("docx empty and multiline paragraph round-trip", () =>
{
    var path = Path.Combine(outputDirectory, "edge-docx-empty-and-text.docx");
    using (var document = new XWPFDocument())
    {
        document.createParagraph();
        var paragraph = document.createParagraph();
        var run = paragraph.createRun();
        run.setText("line1\nline2 <&> \u65e5\u672c\u8a9e");
        run.setBold(true);
        run.setItalic(true);

        using var output = File.Create(path);
        document.write(output);
    }

    using var input = File.OpenRead(path);
    using var read = new XWPFDocument(input);
    AssertEqual(2, read.getParagraphs().Count, "paragraph count");
    AssertEqual("line1\nline2 <&> \u65e5\u672c\u8a9e", read.getParagraphs()[1].getText(), "paragraph text");
});

Probe("pptx empty presentation and blank slide round-trip", () =>
{
    var path = Path.Combine(outputDirectory, "edge-pptx-empty-slide.pptx");
    using (var slideshow = new XMLSlideShow())
    {
        slideshow.createSlide();
        using var output = File.Create(path);
        slideshow.write(output);
    }

    using var input = File.OpenRead(path);
    using var read = new XMLSlideShow(input);
    AssertEqual(1, read.getSlides().Count, "slide count");
    AssertEqual(0, read.getPictureData().Count, "picture count");
});

ExpectException<ArgumentException>("negative row is rejected", () =>
{
    using var workbook = new XSSFWorkbook();
    workbook.createSheet("Rows").createRow(-1);
});

ExpectException<ArgumentException>("negative column is rejected", () =>
{
    using var workbook = new XSSFWorkbook();
    workbook.createSheet("Columns").createRow(0).createCell(-1);
});

Probe("POI-style invalid sheet names should be rejected", () =>
{
    using var workbook = new XSSFWorkbook();
    workbook.createSheet("bad/name");
});

foreach (var result in results)
{
    var status = result.Passed ? "PASS" : "FAIL";
    Console.WriteLine($"{status} {result.Name}");
    if (!result.Passed)
        Console.WriteLine($"     {result.Error}");
}

var failureCount = results.Count(r => !r.Passed);
Console.WriteLine();
Console.WriteLine($"Edge probes complete: {results.Count - failureCount} passed, {failureCount} failed.");

if (failureCount > 0)
    Environment.ExitCode = 1;

void Probe(string name, Action action)
{
    try
    {
        action();
        results.Add(new ProbeResult(name, true, null));
    }
    catch (Exception ex)
    {
        results.Add(new ProbeResult(name, false, $"{ex.GetType().Name}: {ex.Message}"));
    }
}

void ExpectException<TException>(string name, Action action)
    where TException : Exception
{
    try
    {
        action();
        results.Add(new ProbeResult(name, false, $"Expected {typeof(TException).Name}, but no exception was thrown."));
    }
    catch (TException)
    {
        results.Add(new ProbeResult(name, true, null));
    }
    catch (Exception ex)
    {
        results.Add(new ProbeResult(name, false, $"Expected {typeof(TException).Name}, got {ex.GetType().Name}: {ex.Message}"));
    }
}

static void AssertEqual<T>(T expected, T actual, string label)
{
    if (!EqualityComparer<T>.Default.Equals(expected, actual))
        throw new InvalidOperationException($"{label}: expected {expected}, got {actual}.");
}

static void AssertNear(double expected, double actual, string label)
{
    if (Math.Abs(expected - actual) > 0.000001)
        throw new InvalidOperationException($"{label}: expected {expected}, got {actual}.");
}

static void AssertNotNull<T>(T? value, string label)
{
    if (value is null)
        throw new InvalidOperationException($"{label}: expected non-null value.");
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

internal sealed record ProbeResult(string Name, bool Passed, string? Error);
