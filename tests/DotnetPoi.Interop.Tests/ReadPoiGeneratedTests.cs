using DotnetPoi.SS.UserModel;
using DotnetPoi.HSSF.UserModel;
using DotnetPoi.XSSF.UserModel;
using DotnetPoi.XWPF.UserModel;
using DotnetPoi.XSLF.UserModel;
using Xunit;

namespace DotnetPoi.Interop.Tests;

public class ReadPoiGeneratedTests
{
    [Fact]
    [Trait("Category", "ReadFromPoi")]
    [Trait("Format", "Legacy")]
    public void Read_HssfWorkbook_GeneratedByPoi()
    {
        var fixturePath = GetFixturePath("phase6-basic.xls");
        Assert.True(File.Exists(fixturePath), "Run the Java WriteForDotnetTest before this C# read test.");

        using var stream = File.OpenRead(fixturePath);
        using var workbook = new HSSFWorkbook(stream);

        var sheet = workbook.getSheet("From POI HSSF");
        Assert.NotNull(sheet);
        var row = sheet!.getRow(0);
        Assert.NotNull(row);
        Assert.Equal("from apache poi hssf", row!.getCell(0)!.getStringCellValue());
        Assert.Equal(123.75, row.getCell(1)!.getNumericCellValue());
        Assert.False(row.getCell(2)!.getBooleanCellValue());
    }

    [Fact]
    [Trait("Category", "ReadFromPoi")]
    [Trait("Format", "Legacy")]
    public void Read_Phase12HssfStyles_GeneratedByPoi()
    {
        // Phase 12 item 4: Direction A — Java POI writes styled .xls → dotnet-poi reads
        var fixturePath = GetFixturePath("phase12-hssf-styles.xls");
        Assert.True(File.Exists(fixturePath), "Run the Java WriteForDotnetTest#writePhase12HssfStyles before this test.");

        using var stream = File.OpenRead(fixturePath);
        using var workbook = new HSSFWorkbook(stream);

        var sheet = workbook.getSheet("Styles");
        Assert.NotNull(sheet);
        var row = sheet!.getRow(0)!;

        // cell0: numeric with bold Calibri 14pt, center, wrap, border bottom, data format 0.00
        var cell0 = row.getCell(0)!;
        Assert.Equal(CellType.Numeric, cell0.getCellType());
        Assert.Equal(42.5, cell0.getNumericCellValue());
        var style0 = (HSSFCellStyle)cell0.getCellStyle();
        Assert.Equal(HorizontalAlignment.Center, style0.getAlignment());
        Assert.True(style0.getWrapText());
        Assert.Equal(BorderStyle.Thin, style0.getBorderBottom());
        var font0 = (HSSFFont)style0.getFont();
        Assert.True(font0.getBold());
        Assert.True(font0.getItalic());
        Assert.Equal("Calibri", font0.getFontName());
        Assert.Equal(14, font0.getFontHeightInPoints());

        // cell1: right aligned, border left medium
        var cell1 = row.getCell(1)!;
        Assert.Equal(CellType.String, cell1.getCellType());
        Assert.Equal("right", cell1.getStringCellValue());
        var style1 = (HSSFCellStyle)cell1.getCellStyle();
        Assert.Equal(HorizontalAlignment.Right, style1.getAlignment());
        Assert.Equal(BorderStyle.Medium, style1.getBorderLeft());

        // cell2: no special style
        var cell2 = row.getCell(2)!;
        Assert.Equal(CellType.String, cell2.getCellType());
        Assert.Equal("no style", cell2.getStringCellValue());
    }

    [Fact]
    [Trait("Category", "ReadFromPoi")]
    [Trait("Format", "Legacy")]
    public void Read_Phase12HssfLayout_GeneratedByPoi()
    {
        // Phase 12 item 4: Direction A — Java POI writes layout .xls → dotnet-poi reads
        var fixturePath = GetFixturePath("phase12-hssf-layout.xls");
        Assert.True(File.Exists(fixturePath), "Run the Java WriteForDotnetTest#writePhase12HssfLayout before this test.");

        using var stream = File.OpenRead(fixturePath);
        using var workbook = new HSSFWorkbook(stream);

        var sheet = workbook.getSheet("Layout");
        Assert.NotNull(sheet);

        // Column widths
        var width0 = sheet!.getColumnWidth(0);
        var width1 = sheet.getColumnWidth(1);
        Assert.True(width0 > 0, "Column 0 width should be set.");
        Assert.True(width1 > 0, "Column 1 width should be set.");
        Assert.True(width1 > width0, "Column 1 should be wider than column 0.");

        // Hidden column
        Assert.True(sheet.isColumnHidden(2), "Column 2 should be hidden.");

        // Row height
        var row0 = sheet.getRow(0);
        Assert.NotNull(row0);
        Assert.True(row0!.getHeight() > 15.0f, "Row 0 should have custom height > 15pt.");

        // Merged region
        var mergedRegions = sheet.getMergedRegions();
        Assert.True(mergedRegions.Count > 0, "Sheet should have at least one merged region.");
        var merged = mergedRegions[0];
        Assert.Equal(3, merged.FirstRow);
        Assert.Equal(3, merged.LastRow);
        Assert.Equal(0, merged.FirstCol);
        Assert.Equal(2, merged.LastCol);
    }

    [Fact]
    [Trait("Category", "ReadFromPoi")]
    [Trait("Format", "Legacy")]
    public void Read_Phase12HssfUnicode_GeneratedByPoi()
    {
        // Phase 12 item 3: Unicode/Japanese sheet names and string cells — Direction A
        var fixturePath = GetFixturePath("phase12-hssf-unicode.xls");
        Assert.True(File.Exists(fixturePath), "Run the Java WriteForDotnetTest#writePhase12HssfUnicode before this C# read test.");

        using var stream = File.OpenRead(fixturePath);
        using var workbook = new HSSFWorkbook(stream);

        Assert.Equal(2, workbook.getNumberOfSheets());

        var sheet1 = workbook.getSheet("日本語");
        Assert.NotNull(sheet1);
        var row0 = sheet1!.getRow(0);
        Assert.NotNull(row0);
        Assert.Equal("テスト文字列", row0!.getCell(0)!.getStringCellValue());
        Assert.Equal("hello 世界", row0.getCell(1)!.getStringCellValue());
        Assert.Equal("こんにちは", row0.getCell(2)!.getStringCellValue());

        var sheet2 = workbook.getSheet("中文测试");
        Assert.NotNull(sheet2);
        Assert.Equal("汉字测试", sheet2!.getRow(0)!.getCell(0)!.getStringCellValue());
    }

    [Fact]
    [Trait("Category", "ReadFromPoi")]
    [Trait("Format", "Legacy")]
    public void Read_Phase12HssfComprehensive_GeneratedByPoi()
    {
        // Phase 12 item 3: Direction A — Java POI writes .xls → dotnet-poi reads
        var fixturePath = GetFixturePath("phase12-hssf-comprehensive.xls");
        Assert.True(File.Exists(fixturePath), "Run the Java WriteForDotnetTest#writePhase12HssfComprehensive before this C# read test.");

        using var stream = File.OpenRead(fixturePath);
        using var workbook = new HSSFWorkbook(stream);

        Assert.Equal(2, workbook.getNumberOfSheets());

        // Sheet 1: all cell types
        var sheet1 = workbook.getSheet("CellTypes");
        Assert.NotNull(sheet1);
        var row0 = sheet1!.getRow(0);
        Assert.NotNull(row0);

        var c0 = row0!.getCell(0)!;
        Assert.Equal(CellType.String, c0.getCellType());
        Assert.Equal("string value", c0.getStringCellValue());

        var c1 = row0.getCell(1)!;
        Assert.Equal(CellType.Numeric, c1.getCellType());
        Assert.Equal(42.5, c1.getNumericCellValue());

        var c2 = row0.getCell(2)!;
        Assert.Equal(CellType.Boolean, c2.getCellType());
        Assert.True(c2.getBooleanCellValue());

        var c3 = row0.getCell(3)!;
        Assert.Equal(CellType.Boolean, c3.getCellType());
        Assert.False(c3.getBooleanCellValue());

        var c4 = row0.getCell(4)!;
        Assert.Equal(CellType.Error, c4.getCellType());
        Assert.Equal("#DIV/0!", c4.getErrorCellString());

        var c5 = row0.getCell(5)!;
        Assert.Equal(CellType.Error, c5.getCellType());
        Assert.Equal("#N/A", c5.getErrorCellString());

        var c6 = row0.getCell(6)!;
        Assert.Equal(CellType.Blank, c6.getCellType());

        // Sheet 2: sparse layout
        var sheet2 = workbook.getSheet("Sparse");
        Assert.NotNull(sheet2);
        Assert.Equal("row0col0", sheet2!.getRow(0)!.getCell(0)!.getStringCellValue());
        Assert.Equal(99.9, sheet2.getRow(5)!.getCell(3)!.getNumericCellValue());
        Assert.Equal("row10", sheet2.getRow(10)!.getCell(0)!.getStringCellValue());
    }

    [Fact]
    [Trait("Category", "ReadFromPoi")]
    [Trait("Format", "OOXML")]
    public void Read_StringAndNumberWorkbook_GeneratedByPoi()
    {
        var fixturePath = GetFixturePath("phase1-basic.xlsx");
        Assert.True(File.Exists(fixturePath), "Run the Java WriteForDotnetTest before this C# read test.");

        using var stream = File.OpenRead(fixturePath);
        using var workbook = new XSSFWorkbook(stream);

        Assert.Equal(2, workbook.getNumberOfSheets());

        var data = workbook.getSheet("From POI");
        Assert.NotNull(data);
        Assert.Equal("from apache poi", data.getRow(0)!.getCell(0)!.getStringCellValue());
        Assert.Equal(123.25, data.getRow(0)!.getCell(1)!.getNumericCellValue());
        Assert.Equal(0.0, data.getRow(0)!.getCell(2)!.getNumericCellValue());
        Assert.Equal("second row", data.getRow(1)!.getCell(0)!.getStringCellValue());

        var second = workbook.getSheetAt(1);
        Assert.Equal(99.0, second.getRow(2)!.getCell(3)!.getNumericCellValue());
    }

    [Fact]
    [Trait("Category", "ReadFromPoi")]
    [Trait("Format", "OOXML")]
    public void Read_FormulaAndBooleanCells_GeneratedByPoi()
    {
        var fixturePath = GetFixturePath("phase7-formulas.xlsx");
        Assert.True(File.Exists(fixturePath), "Run the Java WriteForDotnetTest before this C# read test.");

        using var stream = File.OpenRead(fixturePath);
        using var workbook = new XSSFWorkbook(stream);
        var sheet = workbook.getSheet("Formulas")!;

        // Row 0: base numeric values written by POI
        Assert.Equal(10.0, sheet.getRow(0)!.getCell(0)!.getNumericCellValue());
        Assert.Equal(20.0, sheet.getRow(0)!.getCell(1)!.getNumericCellValue());

        // Row 1: numeric formula SUM → cached value 30
        var formulaNum = sheet.getRow(1)!.getCell(0)!;
        Assert.Equal(CellType.Formula, formulaNum.getCellType());
        Assert.Equal("A1+B1", formulaNum.getCellFormula());
        Assert.Equal(CellType.Numeric, formulaNum.getCachedFormulaResultType());
        Assert.Equal(30.0, formulaNum.getNumericCellValue());

        // Row 2: string formula → cached "hello world"
        var formulaStr = sheet.getRow(2)!.getCell(0)!;
        Assert.Equal(CellType.Formula, formulaStr.getCellType());
        Assert.Equal("\"hello \"&\"world\"", formulaStr.getCellFormula());
        Assert.Equal(CellType.String, formulaStr.getCachedFormulaResultType());
        Assert.Equal("hello world", formulaStr.getStringCellValue());

        // Row 3: boolean cell (not formula)
        var boolCell = sheet.getRow(3)!.getCell(0)!;
        Assert.Equal(CellType.Boolean, boolCell.getCellType());
        Assert.True(boolCell.getBooleanCellValue());

        // Row 4: error formula → cached #DIV/0!
        var errorCell = sheet.getRow(4)!.getCell(0)!;
        Assert.Equal(CellType.Formula, errorCell.getCellType());
        Assert.Equal("1/0", errorCell.getCellFormula());
        Assert.Equal(CellType.Error, errorCell.getCachedFormulaResultType());
        Assert.Equal("#DIV/0!", errorCell.getErrorCellString());
    }

    [Fact]
    [Trait("Category", "ReadFromPoi")]
    [Trait("Format", "OOXML")]
    public void Read_ForceFormulaRecalculation_GeneratedByPoi()
    {
        var fixturePath = GetFixturePath("phase5-step2-recalc.xlsx");
        Assert.True(File.Exists(fixturePath), "Run the Java WriteForDotnetTest before this C# read test.");

        using var stream = File.OpenRead(fixturePath);
        using var workbook = new XSSFWorkbook(stream);

        Assert.True(workbook.getForceFormulaRecalculation());
        var formula = workbook.getSheet("Recalc")!.getRow(0)!.getCell(0)!;
        Assert.Equal(CellType.Formula, formula.getCellType());
        Assert.Equal("B1+C1", formula.getCellFormula());
    }

    [Fact]
    [Trait("Category", "ReadFromPoi")]
    [Trait("Format", "OOXML")]
    public void Read_DocxComprehensive_GeneratedByPoi()
    {
        var fixturePath = GetFixturePath("phase-docx-comprehensive.docx");
        Assert.True(File.Exists(fixturePath), "Run the Java WriteForDotnetTest before this C# read test.");

        using var stream = File.OpenRead(fixturePath);
        using var doc = new XWPFDocument(stream);

        // --- Paragraph 1: plain text ---
        var p1 = doc.getParagraphs()[0];
        Assert.Equal("First paragraph", p1.getText());

        // --- Paragraph 2: bold + normal ---
        var p2 = doc.getParagraphs()[1];
        var runs2 = p2.getRuns();
        Assert.Equal(2, runs2.Count);
        Assert.Equal("Bold", runs2[0].getText(0));
        Assert.Equal(" and normal", runs2[1].getText(0));

        // --- Paragraph 3: italic ---
        var p3 = doc.getParagraphs()[2];
        Assert.Equal("Italic text", p3.getText());

        // --- Table (2x2) ---
        var tables = doc.getTables();
        Assert.Single(tables);
        var table = tables[0];
        Assert.Equal(2, table.getRows().Count);
        Assert.Equal(2, table.getRows()[0].getCells().Count);
        Assert.Equal("A1", table.getRows()[0].getCells()[0].getParagraphs()[0].getText());
        Assert.Equal("B1", table.getRows()[0].getCells()[1].getParagraphs()[0].getText());
        Assert.Equal("A2", table.getRows()[1].getCells()[0].getParagraphs()[0].getText());
        Assert.Equal("B2", table.getRows()[1].getCells()[1].getParagraphs()[0].getText());

        // --- Hyperlink paragraph ---
        var linkPara = doc.getParagraphs()[3];
        Assert.Equal("Click here for Apache POI", linkPara.getText());

        // --- Header ---
        Assert.Equal("Interop Header", doc.getHeaderText());

        // --- Footer ---
        Assert.Equal("Interop Footer", doc.getFooterText());
    }

    [Fact]
    [Trait("Category", "ReadFromPoi")]
    [Trait("Format", "OOXML")]
    public void Read_PptxWithTextBoxes_GeneratedByPoi()
    {
        var fixturePath = GetFixturePath("phase-pptx-comprehensive.pptx");
        Assert.True(File.Exists(fixturePath), "Run the Java WriteForDotnetTest before this C# read test.");

        using var stream = File.OpenRead(fixturePath);
        using var prs = new XMLSlideShow(stream);

        var slides = prs.getSlides();
        Assert.Single(slides);

        var slide = slides[0];
        var autoShapes = slide.getAutoShapes();
        Assert.Single(autoShapes);

        var tb = autoShapes[0];
        var paragraphs = tb.Paragraphs;
        Assert.Equal(3, paragraphs.Count);

        // Paragraph 0 is an empty default paragraph from POI's createTextBox()
        // Paragraph 1: "Bold Title" (bold, 18pt)
        Assert.Equal("Bold Title", paragraphs[1].getPlainText());

        // Paragraph 2: "Italic subtitle" (italic, 14pt)
        Assert.Equal("Italic subtitle", paragraphs[2].getPlainText());
    }

    [Fact]
    [Trait("Category", "ReadFromPoi")]
    [Trait("Format", "OOXML")]
    public void Read_AutoFilterSheet_GeneratedByPoi()
    {
        var fixturePath = GetFixturePath("phase-autofilter.xlsx");
        Assert.True(File.Exists(fixturePath), "Run the Java WriteForDotnetTest before this C# read test.");

        using var stream = File.OpenRead(fixturePath);
        using var workbook = new XSSFWorkbook(stream);

        var sheet = workbook.getSheet("Data")!;
        Assert.NotNull(sheet);

        var autoFilter = sheet.getAutoFilter();
        Assert.NotNull(autoFilter);
        Assert.Equal(0, autoFilter.FirstRow);
        Assert.Equal(2, autoFilter.LastRow);
        Assert.Equal(0, autoFilter.FirstCol);
        Assert.Equal(1, autoFilter.LastCol);

        Assert.Equal("Category", sheet.getRow(0)!.getCell(0)!.getStringCellValue());
        Assert.Equal(100.0, sheet.getRow(1)!.getCell(1)!.getNumericCellValue());
    }

    [Fact]
    [Trait("Category", "ReadFromPoi")]
    [Trait("Format", "OOXML")]
    public void Read_ProtectedSheet_GeneratedByPoi()
    {
        var fixturePath = GetFixturePath("phase-protection.xlsx");
        Assert.True(File.Exists(fixturePath), "Run the Java WriteForDotnetTest before this C# read test.");

        using var stream = File.OpenRead(fixturePath);
        using var workbook = new XSSFWorkbook(stream);

        Assert.True(workbook.isWorkbookProtected());
        var sheet = workbook.getSheet("Data")!;
        Assert.True(sheet.isSheetProtected());
        Assert.Equal("protected cell", sheet.getRow(0)!.getCell(0)!.getStringCellValue());
    }

    [Fact]
    [Trait("Category", "ReadFromPoi")]
    [Trait("Format", "OOXML")]
    public void Read_ActiveSheet_GeneratedByPoi()
    {
        var fixturePath = GetFixturePath("phase-active-sheet.xlsx");
        Assert.True(File.Exists(fixturePath), "Run the Java WriteForDotnetTest before this C# read test.");

        using var stream = File.OpenRead(fixturePath);
        using var workbook = new XSSFWorkbook(stream);

        Assert.Equal(3, workbook.getNumberOfSheets());
        Assert.Equal(1, workbook.getActiveSheetIndex());
        Assert.NotNull(workbook.getSheetAt(1));
    }

    [Fact(Skip = "Apache POI 5.5.1 does not expose an XWPF field creation API, so the from-poi fixture contains placeholders instead of field codes.")]
    [Trait("Category", "ReadFromPoi")]
    [Trait("Format", "OOXML")]
    public void Read_DocxWithFields_GeneratedByPoi()
    {
        var fixturePath = GetFixturePath("phase-docx-fields.docx");
        Assert.True(File.Exists(fixturePath), "Run the Java WriteForDotnetTest before this C# read test.");

        using var stream = File.OpenRead(fixturePath);
        using var doc = new XWPFDocument(stream);

        var paragraphs = doc.getParagraphs();

        // Paragraph 0: "Page " + PAGE field
        var p0Fields = paragraphs[0].getFields();
        Assert.NotEmpty(p0Fields);
        Assert.Contains("PAGE", p0Fields[0].Instruction);

        // Paragraph 1: TOC field
        Assert.Single(paragraphs[1].getFields());
        Assert.Contains("TOC", paragraphs[1].getFields()[0].Instruction);

        // Paragraph 2: text + MERGEFIELD
        Assert.Equal("Hello ", paragraphs[2].getRuns()[0].getText(0));
        var p2Fields = paragraphs[2].getFields();
        Assert.NotEmpty(p2Fields);
        Assert.Contains("MERGEFIELD", p2Fields[0].Instruction);
    }

    [Fact(Skip = "Fixtures generated by Java tests. Run Java WriteForDotnetTest#writePhase13SampleDoc first.")]
    [Trait("Category", "ReadFromPoi")]
    [Trait("Format", "Legacy")]
    public void Read_Phase13SampleDoc_GeneratedByPoi()
    {
        // Phase 13: Direction A — Java POI writes SampleDoc.doc → dotnet-poi reads
        var fixturePath = GetFixturePath("phase13-sample-doc.doc");
        Assert.True(File.Exists(fixturePath), "Run the Java WriteForDotnetTest#writePhase13SampleDoc before this C# read test.");

        using var stream = File.OpenRead(fixturePath);
        using var doc = new DotnetPoi.HWPF.UserModel.HWPFDocument(stream);

        var text = doc.getText();
        Assert.NotNull(text);
        Assert.True(text.Length > 0, "Document text should be non-empty.");

        // FIB must be readable
        var fib = doc.getFileInformationBlock();
        Assert.True(fib.CcpText >= 0);
        Assert.True(fib.DeclaredTableStreamName is "0Table" or "1Table");

        // Range must compose
        var range = doc.getRange();
        Assert.Equal(text, range.text());
        Assert.True(range.numParagraphs() > 0);

        // Paragraph and run composition
        for (int i = 0; i < range.numParagraphs(); i++)
        {
            var para = range.getParagraph(i);
            var runText = string.Concat(Enumerable.Range(0, para.numCharacterRuns())
                .Select(j => para.getCharacterRun(j).text()));
            Assert.Equal(para.text(), runText);
        }
    }

    private static string GetFixturePath(string fileName)
    {
        var directory = AppContext.BaseDirectory;
        while (directory is not null)
        {
            var fixtures = Path.Combine(directory, "tests", "DotnetPoi.Interop.Tests", "fixtures");
            if (Directory.Exists(fixtures))
            {
                return Path.Combine(fixtures, "from-poi", fileName);
            }

            directory = Directory.GetParent(directory)?.FullName;
        }

        throw new DirectoryNotFoundException("Could not locate tests/DotnetPoi.Interop.Tests/fixtures.");
    }
}
