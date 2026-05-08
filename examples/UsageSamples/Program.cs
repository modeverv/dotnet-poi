using DotnetPoi.SS.UserModel;
using DotnetPoi.SS.Util;
using DotnetPoi.XSSF.UserModel;
using DotnetPoi.XSLF.UserModel;
using DotnetPoi.XWPF.UserModel;
using DotnetPoi.HSSF.UserModel;
using DotnetPoi.HWPF.UserModel;
using DotnetPoi.HSLF.UserModel;
using System.IO.Compression;

var repoRoot = FindRepositoryRoot(AppContext.BaseDirectory);
var outputDirectory = Path.Combine(repoRoot, "examples", "output");
Directory.CreateDirectory(outputDirectory);

var sampleImagePath = Path.Combine(repoRoot, "tests", "test-files", "image.jpg");
var sampleImageBytes = File.ReadAllBytes(sampleImagePath);

var spreadsheetPath = Path.Combine(outputDirectory, "usage-workbook.xlsx");
var documentPath = Path.Combine(outputDirectory, "usage-document.docx");
var presentationPath = Path.Combine(outputDirectory, "usage-presentation.pptx");
var macroWorkbookPath = Path.Combine(outputDirectory, "usage-macro-preserve.xlsm");
var macroTemplatePath = Path.Combine(repoRoot, "tests", "test-files", "example.xlsm");

var legacyXlsPath = Path.Combine(outputDirectory, "usage-workbook.xls");
var legacyDocPath = Path.Combine(outputDirectory, "usage-document.doc");
var legacyDocTemplatePath = Path.Combine(repoRoot, "poi", "test-data", "document", "SampleDoc.doc");
var legacyPptTemplatePath = Path.Combine(repoRoot, "poi", "test-data", "slideshow", "SampleShow.ppt");

CreateSpreadsheet(spreadsheetPath);
CreateMacroWorkbookRoundTrip(macroTemplatePath, macroWorkbookPath);
CreateDocument(documentPath, sampleImageBytes);
CreatePresentation(presentationPath, sampleImageBytes);

CreateLegacyXls(legacyXlsPath);
CreateLegacyDocRoundTrip(legacyDocTemplatePath, legacyDocPath);
ProbeLegacyPpt(legacyPptTemplatePath);

Console.WriteLine("Usage samples generated and verified:");
Console.WriteLine($"  {spreadsheetPath}");
Console.WriteLine($"  {macroWorkbookPath}");
Console.WriteLine($"  {documentPath}");
Console.WriteLine($"  {presentationPath}");
Console.WriteLine($"  {legacyXlsPath}");
Console.WriteLine($"  {legacyDocPath}");

static void CreateSpreadsheet(string outputPath)
{
    using var workbook = new XSSFWorkbook();
    var sheet = workbook.createSheet("Invoice");

    var titleStyle = workbook.createCellStyle();
    var titleFont = workbook.createFont();
    titleFont.setBold(true);
    titleFont.setFontHeightInPoints(16);
    titleStyle.setFont(titleFont);
    titleStyle.setAlignment(HorizontalAlignment.Center);

    var headerStyle = workbook.createCellStyle();
    var headerFont = workbook.createFont();
    headerFont.setBold(true);
    headerStyle.setFont(headerFont);
    headerStyle.setFillForegroundColor((short)IndexedColors.Grey25Percent);
    headerStyle.setFillPattern(FillPatternType.SolidForeground);
    headerStyle.setBorderBottom(BorderStyle.Thin);

    var moneyStyle = workbook.createCellStyle();
    moneyStyle.setDataFormat(workbook.createDataFormat().getFormat("#,##0.00"));

    var totalStyle = workbook.createCellStyle();
    totalStyle.setDataFormat(workbook.createDataFormat().getFormat("#,##0.00"));
    totalStyle.setBorderTop(BorderStyle.Thin);
    totalStyle.setBorderBottom(BorderStyle.Thin);

    var titleRow = sheet.createRow(0);
    var titleCell = titleRow.createCell(0);
    titleCell.setCellValue("Invoice");
    titleCell.setCellStyle(titleStyle);
    sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 3));

    var header = sheet.createRow(2);
    string[] labels = ["Item", "Unit Price", "Quantity", "Line Total"];
    for (var i = 0; i < labels.Length; i++)
    {
        var cell = header.createCell(i);
        cell.setCellValue(labels[i]);
        cell.setCellStyle(headerStyle);
    }

    var rows = new (string Item, double Price, int Quantity)[]
    {
        ("Notebook", 7.5, 4),
        ("Pencil", 1.2, 10),
        ("Folder", 2.8, 6),
    };

    for (var i = 0; i < rows.Length; i++)
    {
        var dataRow = sheet.createRow(i + 3);
        dataRow.createCell(0).setCellValue(rows[i].Item);

        var priceCell = dataRow.createCell(1);
        priceCell.setCellValue(rows[i].Price);
        priceCell.setCellStyle(moneyStyle);

        dataRow.createCell(2).setCellValue(rows[i].Quantity);

        var totalCell = dataRow.createCell(3);
        totalCell.setCellValue(rows[i].Price * rows[i].Quantity);
        totalCell.setCellStyle(totalStyle);
    }

    var rich = new XSSFRichTextString();
    rich.addRun("Note: ", bold: true, italic: false, underline: false, strikethrough: false);
    rich.addRun("quantities must be whole numbers.", bold: false, italic: true, underline: false, strikethrough: false);
    sheet.createRow(7).createCell(0).setCellValue(rich);

    sheet.setColumnWidth(0, 24 * 256);
    sheet.setColumnWidth(1, 14 * 256);
    sheet.setColumnWidth(2, 12 * 256);
    sheet.setColumnWidth(3, 14 * 256);
    sheet.createFreezePane(0, 3);
    sheet.setAutoFilter(new CellRangeAddress(2, 5, 0, 3));

    sheet.AddDataValidation(new XSSFDataValidation
    {
        Sqref = "C4:C6",
        Type = DataValidationType.Whole,
        Operator = DataValidationOperator.Between,
        Formula1 = "1",
        Formula2 = "100",
        PromptTitle = "Quantity",
        PromptMessage = "Enter a whole number from 1 to 100.",
        ErrorTitle = "Invalid quantity",
        ErrorMessage = "Quantity must be between 1 and 100."
    });

    var conditional = new XSSFConditionalFormatting { Sqref = "D4:D6" };
    conditional.Rules.Add(new XSSFCFRule
    {
        Type = ConditionalFormatType.CellIs,
        Operator = "greaterThan",
        Priority = 1,
        DxfId = -1
    });
    conditional.Rules[0].Formulas.Add("20");
    sheet.AddConditionalFormatting(conditional);

    sheet.protectSheet(true);
    workbook.protectWorkbook(true);

    var pivotSheet = workbook.createSheet("Summary");
    var pivotTable = pivotSheet.createPivotTable("A1", "A3:D6", "Invoice");
    pivotTable.RowLabels.Add(0);
    pivotTable.DataColumns.Add(3);

    using (var stream = File.Create(outputPath))
    {
        workbook.write(stream);
    }

    using var readStream = File.OpenRead(outputPath);
    using var loaded = new XSSFWorkbook(readStream);
    var loadedSheet = loaded.getSheet("Invoice")!;
    AssertEqual("Invoice", loadedSheet.getRow(0)!.getCell(0)!.getStringCellValue(), "spreadsheet title");
    AssertEqual(4.0, loadedSheet.getRow(3)!.getCell(2)!.getNumericCellValue(), "spreadsheet quantity");
    AssertEqual(30.0, loadedSheet.getRow(3)!.getCell(3)!.getNumericCellValue(), "spreadsheet total");
    AssertEqual("Note: quantities must be whole numbers.",
        loadedSheet.getRow(7)!.getCell(0)!.getRichStringCellValue().getString(),
        "spreadsheet rich text");
    AssertEqual(true, loadedSheet.isSheetProtected(), "spreadsheet sheet protection");
    AssertEqual(true, loaded.isWorkbookProtected(), "spreadsheet workbook protection");
    AssertEqual("A3:D6", loadedSheet.getAutoFilter()!.FormatAsString(), "spreadsheet auto filter");
    AssertEqual(2, loaded.getNumberOfSheets(), "spreadsheet sheet count including pivot sheet");
}

static void CreateMacroWorkbookRoundTrip(string templatePath, string outputPath)
{
    using (var template = File.OpenRead(templatePath))
    using (var workbook = new XSSFWorkbook(template))
    {
        if (!workbook.HasMacros)
            throw new InvalidOperationException("Macro template should be loaded as an xlsm workbook.");

        var sheet = workbook.getSheetAt(0);
        var row = sheet.getRow(0) ?? sheet.createRow(0);
        row.createCell(5).setCellValue("edited by UsageSamples");

        using var output = File.Create(outputPath);
        workbook.write(output);
    }

    var originalVba = ReadZipEntry(templatePath, "xl/vbaProject.bin");
    var roundTrippedVba = ReadZipEntry(outputPath, "xl/vbaProject.bin");
    AssertBytesEqual(originalVba, roundTrippedVba, "xlsm VBA project");

    using var readBack = File.OpenRead(outputPath);
    using var loaded = new XSSFWorkbook(readBack);
    AssertEqual(true, loaded.HasMacros, "xlsm macro flag");
    AssertEqual("edited by UsageSamples", loaded.getSheetAt(0).getRow(0)!.getCell(5)!.getStringCellValue(), "xlsm edited cell");
}

static void CreateDocument(string outputPath, byte[] imageBytes)
{
    using var document = new XWPFDocument();
    document.setPageSize(16838, 11906);
    document.setLandscape(true);
    document.setMargins(top: 720, right: 720, bottom: 720, left: 720);
    document.setColumns(count: 2, spacingTwips: 720);
    document.setHeaderText("Usage Sample Header");
    document.setFirstHeaderText("Usage Sample First Page Header");
    document.setEvenHeaderText("Usage Sample Even Page Header");
    document.setFooterText("Usage Sample Footer");
    document.setFirstFooterText("Usage Sample First Page Footer");
    document.setEvenFooterText("Usage Sample Even Page Footer");

    var heading = document.createParagraph();
    var headingRun = heading.createRun();
    headingRun.setText("Usage Sample Report");
    headingRun.setBold(true);
    headingRun.setFontSize(18);

    var intro = document.createParagraph();
    intro.createRun().setText("This document was created, saved, and read back with dotnet-poi.");

    var pageField = document.createParagraph();
    pageField.createRun().setText("Page ");
    pageField.addField(" PAGE ", "1");

    var tocField = document.createParagraph();
    tocField.addField("TOC \\o \"1-3\" \\h \\z \\u");

    var mergeField = document.createParagraph();
    mergeField.createRun().setText("Customer: ");
    mergeField.addField("MERGEFIELD CustomerName", "Acme Inc.");

    var link = document.createParagraph();
    var linkRun = link.createRun();
    linkRun.setText("Project website");
    linkRun.setColor("0563C1");
    linkRun.setUnderline(true);
    linkRun.setHyperlink("https://github.com/");

    var table = document.createTable();
    table.addGridCol(3600);
    table.addGridCol(3600);
    AddDocxTableRow(table, "Format", "Status");
    AddDocxTableRow(table, "xlsx", "write/read");
    AddDocxTableRow(table, "docx", "paragraphs/tables/images");

    var imageParagraph = document.createParagraph();
    var picture = imageParagraph.createRun().addPicture(
        imageBytes,
        XWPFPictureData.PICTURE_TYPE_JPEG,
        "sample.jpg",
        width: 1_828_800,
        height: 1_371_600);
    picture.setRotation(15);

    using (var stream = File.Create(outputPath))
    {
        document.write(stream);
    }

    using var readStream = File.OpenRead(outputPath);
    using var loaded = new XWPFDocument(readStream);
    AssertEqual("Usage Sample Report", loaded.getParagraphs()[0].getText(), "document heading");
    AssertEqual("Status", loaded.getTables()[0].getRows()[0].getCells()[1].getParagraphs()[0].getText(), "document table");
    AssertEqual(1, loaded.getAllPictures().Count, "document image count");
    AssertEqual(true, loaded.isLandscape(), "document landscape setting");
    AssertEqual(2, loaded.getColumnCount(), "document column count");
    AssertEqual("Usage Sample Header", loaded.getHeaderText(), "document default header");
    AssertEqual("Usage Sample Footer", loaded.getFooterText(), "document default footer");
    AssertEqual(" PAGE ", loaded.getParagraphs()[2].getFields()[0].Instruction, "document page field");
    AssertEqual("MERGEFIELD CustomerName", loaded.getParagraphs()[4].getFields()[0].Instruction, "document merge field");
}

static void AddDocxTableRow(XWPFTable table, string left, string right)
{
    var row = table.createRow();
    row.createCell().addParagraph().createRun().setText(left);
    row.createCell().addParagraph().createRun().setText(right);
}

static void CreatePresentation(string outputPath, byte[] imageBytes)
{
    using var presentation = new XMLSlideShow();
    var imageIndex = presentation.addPicture(imageBytes, XSLFPictureData.PICTURE_TYPE_JPEG);

    var titleSlide = presentation.createSlide();
    var titleBox = titleSlide.createTextBox();
    titleBox.setAnchor(685_800, 457_200, 7_315_200, 914_400);
    var title = titleBox.addParagraph().addRun("dotnet-poi usage sample");
    title.Bold = true;
    title.FontSize = 28;

    var subtitleBox = titleSlide.createTextBox();
    subtitleBox.setAnchor(685_800, 1_371_600, 7_315_200, 914_400);
    subtitleBox.addParagraph().addRun("Text, images, and tables in a generated pptx.");

    var pictureSlide = presentation.createSlide();
    var picture = presentation.createPicture(pictureSlide, imageIndex);
    picture.setAnchor(914_400, 685_800, 4_572_000, 3_429_000);
    picture.setRotation(8);

    var table = pictureSlide.createTable();
    table.setAnchor(5_715_000, 914_400, 2_743_200, 1_371_600);
    table.addGridCol(1_371_600);
    table.addGridCol(1_371_600);
    AddPptxTableRow(table, "File", "Created");
    AddPptxTableRow(table, "pptx", "yes");

    using (var stream = File.Create(outputPath))
    {
        presentation.write(stream);
    }

    using var readStream = File.OpenRead(outputPath);
    using var loaded = new XMLSlideShow(readStream);
    AssertEqual(2, loaded.getSlides().Count, "presentation slide count");
    AssertEqual("dotnet-poi usage sample",
        loaded.getSlides()[0].getAutoShapes()[0].Paragraphs[0].getPlainText(),
        "presentation title");
    AssertEqual("pptx",
        loaded.getSlides()[1].getTables()[0].Rows[1].Cells[0].Paragraphs[0].getPlainText(),
        "presentation table");
}

static void CreateLegacyXls(string outputPath)
{
    using var workbook = new HSSFWorkbook();
    var sheet = workbook.createSheet("LegacySheet");

    var style = workbook.createCellStyle();
    var font = workbook.createFont();
    font.setBold(true);
    font.setColor((short)IndexedColors.Red);
    style.setFont(font);

    var row = sheet.createRow(0);
    var cell = row.createCell(0);
    cell.setCellValue("Legacy XLS (HSSF)");
    cell.setCellStyle(style);

    row.createCell(1).setCellValue(123.456);

    using (var stream = File.Create(outputPath))
    {
        workbook.write(stream);
    }

    using var readStream = File.OpenRead(outputPath);
    using var loaded = new HSSFWorkbook(readStream);
    AssertEqual("Legacy XLS (HSSF)", loaded.getSheetAt(0).getRow(0).getCell(0).getStringCellValue(), "xls title");
    AssertEqual(123.456, loaded.getSheetAt(0).getRow(0).getCell(1).getNumericCellValue(), "xls value");
}

static void CreateLegacyDocRoundTrip(string templatePath, string outputPath)
{
    if (!File.Exists(templatePath)) return;

    using (var stream = File.OpenRead(templatePath))
    using (var doc = new HWPFDocument(stream))
    {
        doc.appendParagraph("\nEdited by dotnet-poi UsageSamples.");
        doc.replaceText("Sample", "Legacy Sample");

        using var output = File.Create(outputPath);
        doc.write(output);
    }

    using var readBack = File.OpenRead(outputPath);
    using var loaded = new HWPFDocument(readBack);
    var text = loaded.getText();
    if (!text.Contains("Edited by dotnet-poi UsageSamples."))
        throw new InvalidOperationException("doc round-trip: appended text missing.");
}

static void ProbeLegacyPpt(string templatePath)
{
    if (!File.Exists(templatePath)) return;

    using var stream = File.OpenRead(templatePath);
    using var ppt = new HSLFSlideShow(stream);
    var slides = ppt.getSlides();
    if (slides.Count == 0)
        throw new InvalidOperationException("ppt probe: no slides detected.");
}

static void AddPptxTableRow(XSLFTable table, string left, string right)
{
    var row = table.createRow();
    row.createCell().addParagraph().addRun(left);
    row.createCell().addParagraph().addRun(right);
}

static void AssertEqual<T>(T expected, T actual, string label)
{
    if (!EqualityComparer<T>.Default.Equals(expected, actual))
        throw new InvalidOperationException($"{label}: expected {expected}, got {actual}.");
}

static byte[] ReadZipEntry(string packagePath, string entryName)
{
    using var archive = ZipFile.OpenRead(packagePath);
    var entry = archive.GetEntry(entryName)
        ?? throw new InvalidDataException($"{entryName} was not found in {packagePath}.");
    using var stream = entry.Open();
    using var memory = new MemoryStream();
    stream.CopyTo(memory);
    return memory.ToArray();
}

static void AssertBytesEqual(byte[] expected, byte[] actual, string label)
{
    if (!expected.SequenceEqual(actual))
        throw new InvalidOperationException($"{label}: byte content changed.");
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
