using DotnetPoi.XSSF.UserModel;

var repoRoot = FindRepositoryRoot(AppContext.BaseDirectory);
var outputDirectory = Path.Combine(repoRoot, "examples", "output");
Directory.CreateDirectory(outputDirectory);

var outputPath = Path.Combine(outputDirectory, "phase2_5-images-example.xlsx");
WriteWorkbookWithImage(outputPath);
ReadAndAssertWorkbook(outputPath);

Console.WriteLine("Phase 2.5 image example passed.");
Console.WriteLine($"dotnet-poi wrote: {outputPath}");

static void WriteWorkbookWithImage(string outputPath)
{
    using var workbook = new XSSFWorkbook();
    var sheet = workbook.createSheet("Images");
    sheet.createRow(0).createCell(0).setCellValue("embedded png");

    var pictureIndex = workbook.addPicture(OneByOnePng(), XSSFWorkbook.PICTURE_TYPE_PNG);
    var anchor = workbook.getCreationHelper().createClientAnchor();
    anchor.setCol1(1);
    anchor.setRow1(1);
    anchor.setCol2(3);
    anchor.setRow2(4);

    sheet.createDrawingPatriarch().createPicture(anchor, pictureIndex);

    using var stream = File.Create(outputPath);
    workbook.write(stream);
}

static void ReadAndAssertWorkbook(string path)
{
    using var stream = File.OpenRead(path);
    using var workbook = new XSSFWorkbook(stream);

    AssertInt(1, workbook.getNumberOfSheets(), "sheet count");
    AssertString("embedded png", workbook.getSheet("Images")!.getRow(0)!.getCell(0)!.getStringCellValue(), "A1");

    var picture = Single(workbook.getAllPictures(), "picture count");
    AssertInt(XSSFWorkbook.PICTURE_TYPE_PNG, picture.getPictureType(), "picture type");
    AssertString("png", picture.suggestFileExtension(), "picture extension");
    AssertString("image/png", picture.getMimeType(), "picture MIME type");
}

static byte[] OneByOnePng()
{
    return Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO2O8WcAAAAASUVORK5CYII=");
}

static T Single<T>(IReadOnlyList<T> values, string label)
{
    if (values.Count != 1)
    {
        throw new InvalidOperationException($"{label}: expected 1 item, got {values.Count}.");
    }

    return values[0];
}

static void AssertString(string expected, string actual, string label)
{
    if (!string.Equals(expected, actual, StringComparison.Ordinal))
    {
        throw new InvalidOperationException($"{label}: expected '{expected}', got '{actual}'.");
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
