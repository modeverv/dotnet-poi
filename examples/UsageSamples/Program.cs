using DotnetPoi.SS.UserModel;
using DotnetPoi.SS.Util;
using DotnetPoi.XSSF.UserModel;
using DotnetPoi.XSLF.UserModel;
using DotnetPoi.XWPF.UserModel;

var repoRoot = FindRepositoryRoot(AppContext.BaseDirectory);
var outputDirectory = Path.Combine(repoRoot, "examples", "output");
Directory.CreateDirectory(outputDirectory);

var sampleImagePath = Path.Combine(repoRoot, "tests", "test-files", "image.jpg");
var sampleImageBytes = File.ReadAllBytes(sampleImagePath);

var spreadsheetPath = Path.Combine(outputDirectory, "usage-workbook.xlsx");
var documentPath = Path.Combine(outputDirectory, "usage-document.docx");
var presentationPath = Path.Combine(outputDirectory, "usage-presentation.pptx");

CreateSpreadsheet(spreadsheetPath);
CreateDocument(documentPath, sampleImageBytes);
CreatePresentation(presentationPath, sampleImageBytes);

Console.WriteLine("Usage samples generated and verified:");
Console.WriteLine($"  {spreadsheetPath}");
Console.WriteLine($"  {documentPath}");
Console.WriteLine($"  {presentationPath}");

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
}

static void CreateDocument(string outputPath, byte[] imageBytes)
{
    using var document = new XWPFDocument();

    var heading = document.createParagraph();
    var headingRun = heading.createRun();
    headingRun.setText("Usage Sample Report");
    headingRun.setBold(true);
    headingRun.setFontSize(18);

    var intro = document.createParagraph();
    intro.createRun().setText("This document was created, saved, and read back with dotnet-poi.");

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
