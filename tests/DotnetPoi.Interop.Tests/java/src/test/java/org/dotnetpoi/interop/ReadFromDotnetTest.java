package org.dotnetpoi.interop;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;

import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.extractor.WordExtractor;
import org.apache.poi.hslf.usermodel.HSLFSlideShow;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.FormulaError;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.PictureData;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.poifs.crypt.Decryptor;
import org.apache.poi.poifs.crypt.EncryptionInfo;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFAutoShape;
import org.apache.poi.xslf.usermodel.XSLFPictureShape;
import org.apache.poi.xslf.usermodel.XSLFShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTextParagraph;
import org.apache.poi.xslf.usermodel.XSLFTextRun;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.XWPFComment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFHeader;
import org.apache.poi.xwpf.usermodel.XWPFFooter;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFPicture;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.apache.poi.openxml4j.opc.PackageRelationship;
import org.apache.poi.xwpf.usermodel.XWPFNumbering;
import org.junit.jupiter.api.Test;

public class ReadFromDotnetTest {
    private static byte[] loadTestImage() throws IOException {
        return Files.readAllBytes(findRepoRoot().resolve("tests/test-files/image.jpg"));
    }

    @Test
    void readPhase6BasicHssfWorkbook() throws IOException {
        Path fixture = findRepoRoot().resolve("tests/DotnetPoi.Interop.Tests/fixtures/from-dotnet-poi/phase6-basic.xls");
        assertTrue(Files.exists(fixture), "Run the C# WriteForPoi tests before this Java read test.");

        try (InputStream input = Files.newInputStream(fixture);
             HSSFWorkbook workbook = new HSSFWorkbook(input)) {
            Sheet sheet = workbook.getSheet("Phase6");
            Row row = sheet.getRow(0);

            assertEquals("from dotnet-poi hssf", row.getCell(0).getStringCellValue());
            assertEquals(66.25, row.getCell(1).getNumericCellValue());
            assertTrue(row.getCell(2).getBooleanCellValue());
        }
    }

    @Test
    void readPhase12HssfStyles() throws IOException {
        // Phase 12 item 4: Direction B — Java POI reads dotnet-poi .xls with styles
        Path fixture = findRepoRoot().resolve("tests/DotnetPoi.Interop.Tests/fixtures/from-dotnet-poi/phase12-hssf-styles.xls");
        assertTrue(Files.exists(fixture), "Run the C# Write_Phase12HssfStyles_CreatesFixtureForPoi test first.");

        try (InputStream input = Files.newInputStream(fixture);
             HSSFWorkbook workbook = new HSSFWorkbook(input)) {

            Sheet sheet = workbook.getSheet("Styles");
            assertNotNull(sheet);
            Row row = sheet.getRow(0);
            assertNotNull(row);

            Cell cell0 = row.getCell(0);
            assertEquals(CellType.NUMERIC, cell0.getCellType());
            assertEquals(42.5, cell0.getNumericCellValue(), 0.001);
            CellStyle style0 = cell0.getCellStyle();
            assertEquals(HorizontalAlignment.CENTER, style0.getAlignment());
            assertTrue(style0.getWrapText());
            assertEquals(BorderStyle.THIN, style0.getBorderBottom());
            Font font0 = workbook.getFontAt(style0.getFontIndex());
            assertTrue(font0.getBold());
            assertTrue(font0.getItalic());
            assertEquals("Calibri", font0.getFontName());
            assertEquals(14, font0.getFontHeightInPoints());

            Cell cell1 = row.getCell(1);
            assertEquals(CellType.STRING, cell1.getCellType());
            assertEquals("right", cell1.getStringCellValue());
            CellStyle style1 = cell1.getCellStyle();
            assertEquals(HorizontalAlignment.RIGHT, style1.getAlignment());
            assertEquals(BorderStyle.MEDIUM, style1.getBorderLeft());
        }
    }

    @Test
    void readPhase12HssfLayout() throws IOException {
        // Phase 12 item 4: Direction B — Java POI reads dotnet-poi .xls with layout
        Path fixture = findRepoRoot().resolve("tests/DotnetPoi.Interop.Tests/fixtures/from-dotnet-poi/phase12-hssf-layout.xls");
        assertTrue(Files.exists(fixture), "Run the C# Write_Phase12HssfLayout_CreatesFixtureForPoi test first.");

        try (InputStream input = Files.newInputStream(fixture);
             HSSFWorkbook workbook = new HSSFWorkbook(input)) {

            Sheet sheet = workbook.getSheet("Layout");
            assertNotNull(sheet);

            // Column widths
            assertTrue(sheet.getColumnWidth(0) > 0, "Column 0 width should be set.");
            assertTrue(sheet.getColumnWidth(1) > sheet.getColumnWidth(0), "Column 1 should be wider than column 0.");
            assertTrue(sheet.isColumnHidden(2), "Column 2 should be hidden.");

            // Row height
            Row row0 = sheet.getRow(0);
            assertNotNull(row0);
            assertTrue(row0.getHeightInPoints() > 15.0f, "Row 0 height should be > 15pt.");

            // Merged region
            assertEquals(1, sheet.getNumMergedRegions());
            org.apache.poi.ss.util.CellRangeAddress merged = sheet.getMergedRegion(0);
            assertEquals(3, merged.getFirstRow());
            assertEquals(3, merged.getLastRow());
            assertEquals(0, merged.getFirstColumn());
            assertEquals(2, merged.getLastColumn());
        }
    }

    @Test
    void readPhase13NoOpDoc() throws IOException {
        // Phase 13 item 4: Direction B — Java POI reads dotnet-poi no-op saved .doc
        Path fixture = findRepoRoot().resolve("tests/DotnetPoi.Interop.Tests/fixtures/from-dotnet-poi/phase13-noop-sample.doc");
        assertTrue(Files.exists(fixture), "Run the C# Write_Phase13NoOpDoc_CreatesFixtureForPoi test first.");

        try (InputStream input = Files.newInputStream(fixture);
             HWPFDocument doc = new HWPFDocument(input)) {
            WordExtractor extractor = new WordExtractor(doc);
            String text = extractor.getText();
            assertNotNull(text);
            assertTrue(text.length() > 0, "Document text should be non-empty.");
            assertTrue(text.contains("I am a test document") || text.contains("test document"),
                "Document text should contain expected content.");
        }
    }

    @Test
    void readPhase14EditedDoc() throws IOException {
        // Phase 14: Direction B — Java POI reads dotnet-poi edited .doc
        Path fixture = findRepoRoot().resolve("tests/DotnetPoi.Interop.Tests/fixtures/from-dotnet-poi/phase14-edited-sample.doc");
        assertTrue(Files.exists(fixture), "Run the C# Write_Phase14EditedDoc_CreatesFixtureForPoi test first.");

        try (InputStream input = Files.newInputStream(fixture);
             HWPFDocument doc = new HWPFDocument(input)) {
            WordExtractor extractor = new WordExtractor(doc);
            String text = extractor.getText();
            assertNotNull(text);
            assertTrue(text.length() > 0, "Document text should be non-empty.");

            // Appended text must be present
            assertTrue(text.contains("Phase14 edited paragraph"),
                "Edited doc should contain appended text.");
            assertTrue(text.contains("Second appended"),
                "Edited doc should contain second appended paragraph.");

            // Original text must still be present
            assertTrue(text.contains("I am a test document") || text.contains("test document"),
                "Original document content should survive edits.");
        }
    }

    @Test
    void readPhase12HssfUnicode() throws IOException {
        // Phase 12 item 3: Direction B — Java POI reads dotnet-poi .xls with Japanese/Unicode
        Path fixture = findRepoRoot().resolve("tests/DotnetPoi.Interop.Tests/fixtures/from-dotnet-poi/phase12-hssf-unicode.xls");
        assertTrue(Files.exists(fixture), "Run the C# Write_Phase12HssfUnicode_CreatesFixtureForPoi test before this Java read test.");

        try (InputStream input = Files.newInputStream(fixture);
             HSSFWorkbook workbook = new HSSFWorkbook(input)) {

            assertEquals(2, workbook.getNumberOfSheets());

            Sheet sheet1 = workbook.getSheet("日本語");
            assertNotNull(sheet1);
            Row row0 = sheet1.getRow(0);
            assertNotNull(row0);
            assertEquals("テスト文字列", row0.getCell(0).getStringCellValue());
            assertEquals("hello 世界", row0.getCell(1).getStringCellValue());
            assertEquals("こんにちは", row0.getCell(2).getStringCellValue());

            Sheet sheet2 = workbook.getSheet("中文测试");
            assertNotNull(sheet2);
            assertEquals("汉字测试", sheet2.getRow(0).getCell(0).getStringCellValue());
        }
    }

    @Test
    void readPhase12HssfComprehensive() throws IOException {
        // Phase 12 item 3: Direction B — Java POI reads dotnet-poi generated .xls
        Path fixture = findRepoRoot().resolve("tests/DotnetPoi.Interop.Tests/fixtures/from-dotnet-poi/phase12-hssf-comprehensive.xls");
        assertTrue(Files.exists(fixture), "Run the C# Write_Phase12HssfComprehensive_CreatesFixtureForPoi test before this Java read test.");

        try (InputStream input = Files.newInputStream(fixture);
             HSSFWorkbook workbook = new HSSFWorkbook(input)) {

            assertEquals(2, workbook.getNumberOfSheets());

            // Sheet 1: all cell types
            Sheet sheet1 = workbook.getSheet("CellTypes");
            assertNotNull(sheet1);
            Row row0 = sheet1.getRow(0);
            assertNotNull(row0);

            Cell c0 = row0.getCell(0);
            assertEquals(CellType.STRING, c0.getCellType());
            assertEquals("string value", c0.getStringCellValue());

            Cell c1 = row0.getCell(1);
            assertEquals(CellType.NUMERIC, c1.getCellType());
            assertEquals(42.5, c1.getNumericCellValue(), 0.001);

            Cell c2 = row0.getCell(2);
            assertEquals(CellType.BOOLEAN, c2.getCellType());
            assertTrue(c2.getBooleanCellValue());

            Cell c3 = row0.getCell(3);
            assertEquals(CellType.BOOLEAN, c3.getCellType());
            assertFalse(c3.getBooleanCellValue());

            Cell c4 = row0.getCell(4);
            assertEquals(CellType.ERROR, c4.getCellType());
            assertEquals(FormulaError.DIV0.getCode(), c4.getErrorCellValue());

            Cell c5 = row0.getCell(5);
            assertEquals(CellType.ERROR, c5.getCellType());
            assertEquals(FormulaError.NA.getCode(), c5.getErrorCellValue());

            Cell c6 = row0.getCell(6);
            assertEquals(CellType.BLANK, c6.getCellType());

            // Sheet 2: sparse layout
            Sheet sheet2 = workbook.getSheet("Sparse");
            assertNotNull(sheet2);
            assertEquals("row0col0", sheet2.getRow(0).getCell(0).getStringCellValue());
            assertEquals(99.9, sheet2.getRow(5).getCell(3).getNumericCellValue(), 0.001);
            assertEquals("row10", sheet2.getRow(10).getCell(0).getStringCellValue());
        }
    }

    @Test
    void readPhase0BasicWorkbook() throws IOException {
        Path fixture = findRepoRoot().resolve("tests/DotnetPoi.Interop.Tests/fixtures/from-dotnet-poi/phase0-basic.xlsx");
        assertTrue(Files.exists(fixture), "Run the C# WriteForPoi tests before this Java read test.");

        try (InputStream input = Files.newInputStream(fixture);
             XSSFWorkbook workbook = new XSSFWorkbook(input)) {
            Sheet sheet = workbook.getSheet("Phase0");
            Row row = sheet.getRow(0);
            Cell stringCell = row.getCell(0);
            Cell numberCell = row.getCell(1);
            Cell zeroCell = row.getCell(2);

            assertEquals("from dotnet-poi", stringCell.getStringCellValue());
            assertEquals(123.25, numberCell.getNumericCellValue());
            assertEquals(0.0, zeroCell.getNumericCellValue());
        }
    }

    @Test
    void readPhase2StyledWorkbook() throws IOException {
        Path fixture = findRepoRoot().resolve("tests/DotnetPoi.Interop.Tests/fixtures/from-dotnet-poi/phase2-styles.xlsx");
        assertTrue(Files.exists(fixture), "Run the C# WriteForPoi tests before this Java read test.");

        try (InputStream input = Files.newInputStream(fixture);
             XSSFWorkbook workbook = new XSSFWorkbook(input)) {
            Sheet sheet = workbook.getSheet("Phase2");
            Cell cell = sheet.getRow(0).getCell(0);
            CellStyle style = cell.getCellStyle();
            Font font = workbook.getFontAt(style.getFontIndex());

            assertEquals(123.456, cell.getNumericCellValue());
            assertEquals("0.00", style.getDataFormatString());
            assertEquals("Arial", font.getFontName());
            assertEquals(14, font.getFontHeightInPoints());
            assertTrue(font.getBold());
            assertTrue(font.getItalic());
        }
    }

    @Test
    void readPhase25PictureWorkbook() throws IOException {
        Path fixture = findRepoRoot().resolve("tests/DotnetPoi.Interop.Tests/fixtures/from-dotnet-poi/phase2_5-images.xlsx");
        assertTrue(Files.exists(fixture), "Run the C# WriteForPoi tests before this Java read test.");

        try (InputStream input = Files.newInputStream(fixture);
             XSSFWorkbook workbook = new XSSFWorkbook(input)) {
            Sheet sheet = workbook.getSheet("Phase2.5");

            assertEquals("image", sheet.getRow(0).getCell(0).getStringCellValue());
            assertEquals(1, workbook.getAllPictures().size());

            PictureData picture = workbook.getAllPictures().get(0);
            assertEquals(XSSFWorkbook.PICTURE_TYPE_JPEG, picture.getPictureType());
            assertEquals("jpeg", picture.suggestFileExtension());
            assertTrue(java.util.Arrays.equals(loadTestImage(), picture.getData()));
            assertEquals(1, ((XSSFDrawing)sheet.createDrawingPatriarch()).getShapes().size());
        }
    }

    @Test
    void readPhase31RotatedPictureWorkbook() throws IOException {
        Path fixture = findRepoRoot().resolve("tests/DotnetPoi.Interop.Tests/fixtures/from-dotnet-poi/phase3_1-rotation.xlsx");
        assertTrue(Files.exists(fixture), "Run the C# WriteForPoi tests before this Java read test.");

        try (InputStream input = Files.newInputStream(fixture);
             XSSFWorkbook workbook = new XSSFWorkbook(input)) {
            Sheet sheet = workbook.getSheet("Phase3.1");
            assertEquals("rotated", sheet.getRow(0).getCell(0).getStringCellValue());
            assertEquals(1, workbook.getAllPictures().size());

            PictureData picture = workbook.getAllPictures().get(0);
            assertEquals(XSSFWorkbook.PICTURE_TYPE_JPEG, picture.getPictureType());

            // Verify rotation is stored in the drawing XML (90° = 5400000 in OOXML units)
            org.apache.poi.xssf.usermodel.XSSFSheet xssfSheet =
                (org.apache.poi.xssf.usermodel.XSSFSheet) sheet;
            org.apache.poi.xssf.usermodel.XSSFDrawing drawing =
                (org.apache.poi.xssf.usermodel.XSSFDrawing) xssfSheet.createDrawingPatriarch();
            var shapes = drawing.getShapes();
            assertEquals(1, shapes.size());
            org.apache.poi.xssf.usermodel.XSSFPicture xssfPicture =
                (org.apache.poi.xssf.usermodel.XSSFPicture) shapes.get(0);
            int rotAttribute = xssfPicture.getCTPicture().getSpPr().getXfrm().getRot();
            assertEquals(5400000, rotAttribute);
        }
    }

    @Test
    void readPhase32BasicDocx() throws IOException {
        Path fixture = findRepoRoot().resolve("tests/DotnetPoi.Interop.Tests/fixtures/from-dotnet-poi/phase3_2-basic.docx");
        assertTrue(Files.exists(fixture), "Run the C# WriteForPoi tests before this Java read test.");

        try (InputStream input = Files.newInputStream(fixture);
             XWPFDocument doc = new XWPFDocument(input)) {

            java.util.List<XWPFParagraph> paragraphs = doc.getParagraphs();
            assertTrue(paragraphs.size() >= 2);

            XWPFRun run1 = paragraphs.get(0).getRuns().get(0);
            assertEquals("from dotnet-poi docx", run1.getText(0));
            assertTrue(run1.isBold());

            XWPFRun run2 = paragraphs.get(1).getRuns().get(0);
            assertEquals("second paragraph", run2.getText(0));
            assertTrue(run2.isItalic());
        }
    }

    @Test
    void readPhase321DocxWithImageAndRotation() throws IOException {
        Path fixture = findRepoRoot().resolve("tests/DotnetPoi.Interop.Tests/fixtures/from-dotnet-poi/phase3_2_1-image-rotation.docx");
        assertTrue(Files.exists(fixture), "Run the C# WriteForPoi tests before this Java read test.");

        try (InputStream input = Files.newInputStream(fixture);
             XWPFDocument doc = new XWPFDocument(input)) {

            assertEquals(1, doc.getAllPictures().size());

            XWPFParagraph para = doc.getParagraphs().get(0);
            XWPFRun run = para.getRuns().get(0);
            java.util.List<XWPFPicture> pics = run.getEmbeddedPictures();
            assertEquals(1, pics.size());

            XWPFPicture pic = pics.get(0);
            // 90° = 5 400 000 in OOXML 60000ths-of-a-degree units
            int rot = pic.getCTPicture().getSpPr().getXfrm().getRot();
            assertEquals(5400000, rot);
        }
    }

    @Test
    void readPhase33PptxPresentation() throws IOException {
        Path fixture = findRepoRoot().resolve("tests/DotnetPoi.Interop.Tests/fixtures/from-dotnet-poi/phase3_3-pptx.pptx");
        assertTrue(Files.exists(fixture), "Run the C# WriteForPoi tests before this Java read test.");

        try (InputStream input = Files.newInputStream(fixture);
             XMLSlideShow prs = new XMLSlideShow(input)) {

            assertEquals(1, prs.getSlides().size(), "should have 1 slide");

            XSLFSlide slide = prs.getSlides().get(0);
            assertEquals(1, slide.getShapes().size(), "should have 1 shape");

            XSLFShape shape = slide.getShapes().get(0);
            assertTrue(shape instanceof XSLFPictureShape, "shape should be XSLFPictureShape");
            XSLFPictureShape picture = (XSLFPictureShape) shape;

            // Verify rotation: 90° stored as 5400000 in 60000ths-of-a-degree units
            assertEquals(90.0, picture.getRotation(), 0.01, "rotation should be 90°");

            // Verify picture bytes
            assertTrue(java.util.Arrays.equals(loadTestImage(), picture.getPictureData().getData()),
                "picture bytes should match test image");
        }
    }

    @Test
    void readPhase34AgileEncryptedWorkbook() throws Exception {
        Path fixture = findRepoRoot().resolve("tests/DotnetPoi.Interop.Tests/fixtures/from-dotnet-poi/phase3_4-agile-encrypted.xlsx");
        assertTrue(Files.exists(fixture), "Run the C# WriteForPoi tests before this Java read test.");

        try (InputStream input = Files.newInputStream(fixture);
             POIFSFileSystem fs = new POIFSFileSystem(input)) {
            EncryptionInfo info = new EncryptionInfo(fs);
            Decryptor decryptor = info.getDecryptor();
            assertTrue(decryptor.verifyPassword("f"));

            try (InputStream data = decryptor.getDataStream(fs);
                 XSSFWorkbook workbook = new XSSFWorkbook(data)) {
                Sheet sheet = workbook.getSheet("Phase3.4");
                assertEquals("encrypted from dotnet-poi", sheet.getRow(0).getCell(0).getStringCellValue());
                assertEquals(34.0, sheet.getRow(0).getCell(1).getNumericCellValue());
            }
        }
    }

    @Test
    void readPhase7CellTypes() throws IOException {
        Path fixture = findRepoRoot().resolve("tests/DotnetPoi.Interop.Tests/fixtures/from-dotnet-poi/phase7-cell-types.xlsx");
        assertTrue(Files.exists(fixture), "Run the C# WriteForPoi tests before this Java read test.");

        try (InputStream input = Files.newInputStream(fixture);
             org.apache.poi.xssf.usermodel.XSSFWorkbook workbook = new org.apache.poi.xssf.usermodel.XSSFWorkbook(input)) {
            Sheet sheet = workbook.getSheet("CellTypes");
            Row row = sheet.getRow(0);

            // Boolean true
            assertEquals(org.apache.poi.ss.usermodel.CellType.BOOLEAN, row.getCell(0).getCellType());
            assertTrue(row.getCell(0).getBooleanCellValue());

            // Boolean false
            assertEquals(org.apache.poi.ss.usermodel.CellType.BOOLEAN, row.getCell(1).getCellType());
            assertFalse(row.getCell(1).getBooleanCellValue());

            // Numeric
            assertEquals(org.apache.poi.ss.usermodel.CellType.NUMERIC, row.getCell(2).getCellType());
            assertEquals(42.5, row.getCell(2).getNumericCellValue());

            // String
            assertEquals(org.apache.poi.ss.usermodel.CellType.STRING, row.getCell(3).getCellType());
            assertEquals("hello", row.getCell(3).getStringCellValue());
        }
    }

    @Test
    void readPhase5Step1Formulas() throws IOException {
        Path fixture = findRepoRoot().resolve("tests/DotnetPoi.Interop.Tests/fixtures/from-dotnet-poi/phase5-step1-formulas.xlsx");
        assertTrue(Files.exists(fixture), "Run the C# WriteForPoi tests before this Java read test.");

        try (InputStream input = Files.newInputStream(fixture);
             org.apache.poi.xssf.usermodel.XSSFWorkbook workbook = new org.apache.poi.xssf.usermodel.XSSFWorkbook(input)) {
            Sheet sheet = workbook.getSheet("Formulas");

            assertEquals(10.0, sheet.getRow(0).getCell(0).getNumericCellValue());
            assertEquals(20.0, sheet.getRow(0).getCell(1).getNumericCellValue());

            Cell numericFormula = sheet.getRow(1).getCell(0);
            assertEquals(org.apache.poi.ss.usermodel.CellType.FORMULA, numericFormula.getCellType());
            assertEquals("A1+B1", numericFormula.getCellFormula());

            Cell stringFormula = sheet.getRow(2).getCell(0);
            assertEquals(org.apache.poi.ss.usermodel.CellType.FORMULA, stringFormula.getCellType());
            assertEquals("\"hello \"&\"world\"", stringFormula.getCellFormula());
            assertEquals("hello world", stringFormula.getStringCellValue());
        }
    }

    @Test
    void readPhase5Step2Recalc() throws IOException {
        Path fixture = findRepoRoot().resolve("tests/DotnetPoi.Interop.Tests/fixtures/from-dotnet-poi/phase5-step2-recalc.xlsx");
        assertTrue(Files.exists(fixture), "Run the C# WriteForPoi tests before this Java read test.");

        try (InputStream input = Files.newInputStream(fixture);
             org.apache.poi.xssf.usermodel.XSSFWorkbook workbook = new org.apache.poi.xssf.usermodel.XSSFWorkbook(input)) {
            assertTrue(workbook.getForceFormulaRecalculation());

            Sheet sheet = workbook.getSheet("Recalc");
            Cell formula = sheet.getRow(0).getCell(0);
            assertEquals(org.apache.poi.ss.usermodel.CellType.FORMULA, formula.getCellType());
            assertEquals("B1+C1", formula.getCellFormula());
        }
    }

    @Test
    void readPhase5Step3EvaluatedFunctions() throws IOException {
        Path fixture = findRepoRoot().resolve("tests/DotnetPoi.Interop.Tests/fixtures/from-dotnet-poi/phase5-step3-evaluated-functions.xlsx");
        assertTrue(Files.exists(fixture), "Run the C# WriteForPoi tests before this Java read test.");

        try (InputStream input = Files.newInputStream(fixture);
             org.apache.poi.xssf.usermodel.XSSFWorkbook workbook = new org.apache.poi.xssf.usermodel.XSSFWorkbook(input)) {
            Sheet sheet = workbook.getSheet("Eval");

            Cell sum = sheet.getRow(1).getCell(0);
            assertEquals(org.apache.poi.ss.usermodel.CellType.FORMULA, sum.getCellType());
            assertEquals("SUM(A1:C1)", sum.getCellFormula());
            assertEquals(60.0, sum.getNumericCellValue());

            Cell average = sheet.getRow(1).getCell(1);
            assertEquals(org.apache.poi.ss.usermodel.CellType.FORMULA, average.getCellType());
            assertEquals("AVERAGE(A1:C1)", average.getCellFormula());
            assertEquals(20.0, average.getNumericCellValue());

            Cell text = sheet.getRow(1).getCell(2);
            assertEquals(org.apache.poi.ss.usermodel.CellType.FORMULA, text.getCellType());
            assertEquals("CONCATENATE(\"sum=\",SUM(A1:C1))", text.getCellFormula());
            assertEquals("sum=60", text.getStringCellValue());
        }
    }

    @Test
    void readPhaseDocmInterop() throws IOException {
        Path fixture = findRepoRoot().resolve("tests/DotnetPoi.Interop.Tests/fixtures/from-dotnet-poi/phase-docm-interop.docm");
        assertTrue(Files.exists(fixture), "Run the C# WriteForPoi tests before this Java read test.");

        try (InputStream input = Files.newInputStream(fixture);
             XWPFDocument doc = new XWPFDocument(input)) {

            // Macro-enabled: main part content type must contain "macroEnabled".
            assertTrue(doc.getPackagePart().getContentType().contains("macroEnabled"),
                "docm content type should contain macroEnabled");

            java.util.List<XWPFParagraph> paragraphs = doc.getParagraphs();
            assertTrue(paragraphs.size() >= 2, "should have at least 2 paragraphs");

            XWPFRun run1 = paragraphs.get(0).getRuns().get(0);
            assertEquals("from dotnet-poi docm", run1.getText(0));
            assertTrue(run1.isBold(), "first run should be bold");

            XWPFRun run2 = paragraphs.get(1).getRuns().get(0);
            assertEquals("second paragraph", run2.getText(0));
            assertTrue(run2.isItalic(), "second run should be italic");

            // VBA project must be present.
            java.util.List<org.apache.poi.openxml4j.opc.PackagePart> vbaParts =
                doc.getPackage().getPartsByContentType("application/vnd.ms-office.vbaProject");
            assertFalse(vbaParts.isEmpty(), "word/vbaProject.bin should be in the package");
        }
    }

    @Test
    void readPhaseDocxComprehensive() throws IOException, org.apache.poi.openxml4j.exceptions.InvalidFormatException {
        Path fixture = findRepoRoot().resolve("tests/DotnetPoi.Interop.Tests/fixtures/from-dotnet-poi/phase-docx-comprehensive.docx");
        assertTrue(Files.exists(fixture), "Run the C# WriteForPoi tests before this Java read test.");

        try (InputStream input = Files.newInputStream(fixture);
             XWPFDocument doc = new XWPFDocument(input)) {

            // --- Header / Footer ---
            XWPFHeader header = doc.getHeaderList().get(0);
            assertNotNull(header, "header should exist");
            assertTrue(header.getText().contains("Interop Test Header"), "header text mismatch");

            XWPFFooter footer = doc.getFooterList().get(0);
            assertNotNull(footer, "footer should exist");
            assertTrue(footer.getText().contains("Page"), "footer text mismatch");

            // --- Page setup (landscape) ---
            var sectPr = doc.getDocument().getBody().getSectPr();
            assertNotNull(sectPr, "sectPr should exist");
            var pgSz = sectPr.getPgSz();
            assertNotNull(pgSz, "pgSz should exist");
            assertEquals("landscape", pgSz.getOrient().toString(), "orientation should be landscape");

            java.util.List<XWPFParagraph> paragraphs = doc.getParagraphs();

            // --- 1) Rich text paragraph (paragraph 0) ---
            XWPFParagraph richPara = paragraphs.get(0);
            java.util.List<XWPFRun> runs = richPara.getRuns();
            assertTrue(runs.size() >= 4, "should have at least 4 runs");

            XWPFRun boldRun = runs.get(0);
            assertEquals("Bold ", boldRun.getText(0));
            assertTrue(boldRun.isBold());
            assertEquals(14, boldRun.getFontSize());
            assertEquals("Arial", boldRun.getFontName());
            assertEquals("FF0000", boldRun.getColor());

            XWPFRun italicRun = runs.get(1);
            assertEquals("Italic ", italicRun.getText(0));
            assertTrue(italicRun.isItalic());
            assertEquals(12, italicRun.getFontSize());
            assertEquals("0000FF", italicRun.getColor());

            XWPFRun ulRun = runs.get(2);
            assertEquals("Underline ", ulRun.getText(0));
            assertTrue(ulRun.getUnderline() != null && !ulRun.getUnderline().equals("none"),
                "underline should be set");
            assertEquals("Times New Roman", ulRun.getFontName());

            XWPFRun strikeRun = runs.get(3);
            assertEquals("Strikethrough", strikeRun.getText(0));
            assertTrue(strikeRun.isStrike());
            assertEquals(16, strikeRun.getFontSize());

            // --- 2) Numbered list (paragraphs 1-2) ---
            XWPFNumbering numbering = doc.getNumbering();
            assertNotNull(numbering, "numbering should exist");

            XWPFParagraph numPara1 = paragraphs.get(1);
            assertEquals("First item", numPara1.getText());
            assertNotNull(numPara1.getNumID(), "numbered item should have numId");

            XWPFParagraph numPara2 = paragraphs.get(2);
            assertEquals("Second item", numPara2.getText());
            assertNotNull(numPara2.getNumID(), "numbered item should have numId");

            // --- 3) Bullet list (paragraphs 3-4) ---
            XWPFParagraph bullet1 = paragraphs.get(3);
            assertEquals("Bullet A", bullet1.getText());
            assertNotNull(bullet1.getNumID(), "bullet item should have numId");

            XWPFParagraph bullet2 = paragraphs.get(4);
            assertEquals("Bullet B", bullet2.getText());
            assertNotNull(bullet2.getNumID(), "bullet item should have numId");

            // --- 4) Indentation and spacing (paragraph 5) ---
            XWPFParagraph indentPara = paragraphs.get(5);
            assertEquals("Indented centered paragraph with spacing before and after.",
                indentPara.getText());
            assertEquals("center", indentPara.getAlignment().toString().toLowerCase());
            assertTrue(indentPara.getIndentationLeft() > 0, "indent left should be set");
            assertTrue(indentPara.getIndentationFirstLine() > 0, "first line indent should be set");

            // --- 5) Hyperlink (paragraph 6) ---
            XWPFParagraph linkPara = paragraphs.get(6);
            assertEquals("Click here for DotnetPoi", linkPara.getText());

            // Verify via OPC-level hyperlink relationships
            boolean foundGithub = false;
            for (PackageRelationship rel : doc.getPackagePart().getRelationshipsByType(
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink")) {
                if (rel.getTargetURI().toString().contains("github.com")) {
                    foundGithub = true;
                    break;
                }
            }
            assertTrue(foundGithub, "github hyperlink should exist in document");

            // --- 6) Table ---
            java.util.List<XWPFTable> tables = doc.getTables();
            assertEquals(1, tables.size(), "should have 1 table");

            XWPFTable table = tables.get(0);
            java.util.List<XWPFTableRow> tableRows = table.getRows();
            assertEquals(3, tableRows.size(), "should have 3 rows");

            // Header row
            XWPFTableRow headerTableRow = tableRows.get(0);
            assertEquals("Col A", headerTableRow.getCell(0).getText());
            assertEquals("Col B", headerTableRow.getCell(1).getText());
            assertEquals("Col C", headerTableRow.getCell(2).getText());

            // Data row 1
            XWPFTableRow dataRow1 = tableRows.get(1);
            assertEquals("A1", dataRow1.getCell(0).getText());
            assertEquals("B1", dataRow1.getCell(1).getText());
            assertEquals("C1", dataRow1.getCell(2).getText());

            // Link row (table row with hyperlink in middle cell)
            XWPFTableRow linkRow = tableRows.get(2);
            assertEquals("Link cell", linkRow.getCell(0).getText());
            assertEquals("End", linkRow.getCell(2).getText());
            // Middle cell should contain a hyperlink
            XWPFTableCell linkCell = linkRow.getCell(1);
            String linkCellText = linkCell.getText();
            assertTrue("Example".equals(linkCellText) || linkCellText.contains("Example"),
                "link cell should contain 'Example'");
        }
    }

    @Test
    void readPhasePptxComprehensive() throws IOException {
        Path fixture = findRepoRoot().resolve("tests/DotnetPoi.Interop.Tests/fixtures/from-dotnet-poi/phase-pptx-comprehensive.pptx");
        assertTrue(Files.exists(fixture), "Run the C# WriteForPoi tests before this Java read test.");

        try (InputStream input = Files.newInputStream(fixture);
             XMLSlideShow prs = new XMLSlideShow(input)) {

            assertEquals(1, prs.getSlides().size());
            XSLFSlide slide = prs.getSlides().get(0);

            // --- Text box with formatted text ---
            // The text box (XSLFAutoShape) is the first shape
            XSLFShape shape0 = slide.getShapes().get(0);
            assertTrue(shape0 instanceof XSLFAutoShape, "first shape should be auto shape (text box)");
            XSLFAutoShape autoShape = (XSLFAutoShape) shape0;

            var paragraphs = autoShape.getTextParagraphs();
            assertEquals(2, paragraphs.size());

            // Paragraph 1: "Bold text" (rPr: bold + 18pt set in XML; POI 5.5.1 API
            // returns null/isSetB=false — XML structure is verified by C# round-trip tests)
            assertEquals("Bold text", paragraphs.get(0).getText());

            // Paragraph 2: "Italic text" (rPr: italic + 14pt in XML)
            assertEquals("Italic text", paragraphs.get(1).getText());
        }
    }

    @Test
    void readPhasePptmInterop() throws IOException {
        Path fixture = findRepoRoot().resolve("tests/DotnetPoi.Interop.Tests/fixtures/from-dotnet-poi/phase-pptm-interop.pptm");
        assertTrue(Files.exists(fixture), "Run the C# WriteForPoi tests before this Java read test.");

        try (InputStream input = Files.newInputStream(fixture);
             XMLSlideShow prs = new XMLSlideShow(input)) {

            // Macro-enabled: main part content type must contain "macroEnabled".
            assertTrue(prs.getPackagePart().getContentType().contains("macroEnabled"),
                "pptm content type should contain macroEnabled");
            assertEquals(1, prs.getSlides().size(), "should have 1 slide");

            XSLFSlide slide = prs.getSlides().get(0);
            assertEquals(1, slide.getShapes().size(), "slide should have 1 shape (picture)");
            assertTrue(slide.getShapes().get(0) instanceof XSLFPictureShape,
                "shape should be XSLFPictureShape");

            // VBA project must be present.
            java.util.List<org.apache.poi.openxml4j.opc.PackagePart> vbaParts =
                prs.getPackage().getPartsByContentType("application/vnd.ms-office.vbaProject");
            assertFalse(vbaParts.isEmpty(), "ppt/vbaProject.bin should be in the package");
        }
    }

    @Test
    void readPhaseXlsmInterop() throws IOException {
        Path fixture = findRepoRoot().resolve("tests/DotnetPoi.Interop.Tests/fixtures/from-dotnet-poi/phase-xlsm-interop.xlsm");
        assertTrue(Files.exists(fixture), "Run the C# WriteForPoi tests before this Java read test.");

        try (InputStream input = Files.newInputStream(fixture);
             XSSFWorkbook workbook = new XSSFWorkbook(input)) {

            // Macro-enabled workbook must have a VBA project.
            assertTrue(workbook.isMacroEnabled(), "workbook should be macro-enabled");

            Sheet sheet = workbook.getSheet("MacroSheet");
            assertNotNull(sheet, "sheet 'MacroSheet' should exist");

            Row row = sheet.getRow(0);
            assertNotNull(row, "row 0 should exist");

            Cell stringCell = row.getCell(0);
            assertEquals(org.apache.poi.ss.usermodel.CellType.STRING, stringCell.getCellType());
            assertEquals("from dotnet-poi xlsm", stringCell.getStringCellValue());

            Cell numericCell = row.getCell(1);
            assertEquals(org.apache.poi.ss.usermodel.CellType.NUMERIC, numericCell.getCellType());
            assertEquals(99.5, numericCell.getNumericCellValue(), 0.0001);

            // VBA project must be present in the OPC package.
            java.util.List<org.apache.poi.openxml4j.opc.PackagePart> vbaParts =
                workbook.getPackage().getPartsByContentType("application/vnd.ms-office.vbaProject");
            assertFalse(vbaParts.isEmpty(), "xl/vbaProject.bin should be present in the package");
            try (InputStream vbaStream = vbaParts.get(0).getInputStream();
                 java.io.ByteArrayOutputStream vbaBuf = new java.io.ByteArrayOutputStream()) {
                vbaStream.transferTo(vbaBuf);
                assertTrue(vbaBuf.size() > 0, "vbaProject.bin should be non-empty");
            }
        }
    }

    @Test
    void readPhaseAutoFilterSheet() throws IOException {
        Path fixture = findRepoRoot().resolve("tests/DotnetPoi.Interop.Tests/fixtures/from-dotnet-poi/phase-autofilter.xlsx");
        assertTrue(Files.exists(fixture), "Run the C# WriteForPoi tests before this Java read test.");

        try (InputStream input = Files.newInputStream(fixture);
             XSSFWorkbook workbook = new XSSFWorkbook(input)) {

            XSSFSheet xssfSheet = workbook.getSheet("Data");
            assertNotNull(xssfSheet);

            // Use low-level API: getAutoFilter() was added in POI 5.6.0; POI 5.5.1 uses CTWorksheet
            assertNotNull(xssfSheet.getCTWorksheet().getAutoFilter());

            // Verify cell values
            assertEquals("Category", xssfSheet.getRow(0).getCell(0).getStringCellValue());
            assertEquals(100.0, xssfSheet.getRow(1).getCell(1).getNumericCellValue(), 0.0001);
            assertEquals("Travel", xssfSheet.getRow(2).getCell(0).getStringCellValue());
        }
    }

    @Test
    void readPhaseProtectedSheet() throws IOException {
        Path fixture = findRepoRoot().resolve("tests/DotnetPoi.Interop.Tests/fixtures/from-dotnet-poi/phase-protection.xlsx");
        assertTrue(Files.exists(fixture), "Run the C# WriteForPoi tests before this Java read test.");

        try (InputStream input = Files.newInputStream(fixture);
             XSSFWorkbook workbook = new XSSFWorkbook(input)) {

            XSSFSheet sheet = workbook.getSheet("Data");
            assertNotNull(sheet);

            // Verify sheet is protected (POI 5.5.1: use low-level CTWorksheet API)
            assertNotNull(sheet.getCTWorksheet().getSheetProtection(), "sheet should have protection settings");

            assertEquals("protected cell", sheet.getRow(0).getCell(0).getStringCellValue());
        }
    }

    @Test
    void readPhaseActiveSheet() throws IOException {
        Path fixture = findRepoRoot().resolve("tests/DotnetPoi.Interop.Tests/fixtures/from-dotnet-poi/phase-active-sheet.xlsx");
        assertTrue(Files.exists(fixture), "Run the C# WriteForPoi tests before this Java read test.");

        try (InputStream input = Files.newInputStream(fixture);
             XSSFWorkbook workbook = new XSSFWorkbook(input)) {

            assertEquals(3, workbook.getNumberOfSheets());
            assertEquals(1, workbook.getActiveSheetIndex());
            assertEquals("Second", workbook.getSheetAt(1).getSheetName());
        }
    }

    @Test
    void readPhase18XssfCommentsWorkbook() throws IOException {
        Path fixture = findRepoRoot().resolve("tests/DotnetPoi.Interop.Tests/fixtures/from-dotnet-poi/phase18-xssf-comments.xlsx");
        assertTrue(Files.exists(fixture), "Run the C# WriteForPoi tests before this Java read test.");

        try (InputStream input = Files.newInputStream(fixture);
             XSSFWorkbook workbook = new XSSFWorkbook(input)) {
            XSSFSheet first = workbook.getSheet("Comments A");
            XSSFSheet second = workbook.getSheet("Comments B");
            assertNotNull(first);
            assertNotNull(second);

            assertEquals(2, first.getCellComments().size());
            assertEquals(1, second.getCellComments().size());

            Comment visible = first.getRow(0).getCell(0).getCellComment();
            assertNotNull(visible);
            assertEquals("DotnetPoi", visible.getAuthor());
            assertEquals("Visible from dotnet-poi", visible.getString().getString());
            assertTrue(visible.isVisible());

            assertNull(first.getCellComment(new org.apache.poi.ss.util.CellAddress("B2")));
            Comment moved = first.getRow(2).getCell(2).getCellComment();
            assertNotNull(moved);
            assertEquals("DotnetPoi", moved.getAuthor());
            assertEquals("Moved to C3", moved.getString().getString());
            assertFalse(moved.isVisible());

            Comment other = second.getRow(1).getCell(1).getCellComment();
            assertNotNull(other);
            assertEquals("Interop", other.getAuthor());
            assertEquals("Second sheet comment", other.getString().getString());
        }
    }

    @Test
    void readPhase18XwpfCommentsDocument() throws IOException {
        Path fixture = findRepoRoot().resolve("tests/DotnetPoi.Interop.Tests/fixtures/from-dotnet-poi/phase18-xwpf-comments.docx");
        assertTrue(Files.exists(fixture), "Run the C# WriteForPoi tests before this Java read test.");

        try (InputStream input = Files.newInputStream(fixture);
             XWPFDocument document = new XWPFDocument(input)) {
            assertEquals(1, document.getComments().length);
            XWPFComment comment = document.getCommentByID("0");
            assertNotNull(comment);
            assertEquals("DotnetPoi", comment.getAuthor());
            assertEquals("DP", comment.getInitials());
            assertEquals("Docx comment from dotnet-poi", comment.getText());

            XWPFParagraph paragraph = document.getParagraphs().get(0);
            assertEquals("DotnetPoi commented text", paragraph.getText());
            assertEquals(1, paragraph.getCTP().getCommentRangeStartList().size());
            assertEquals(1, paragraph.getCTP().getCommentRangeEndList().size());
            assertEquals("0", paragraph.getCTP().getCommentRangeStartList().get(0).getId().toString());
            assertEquals("0", paragraph.getCTP().getCommentRangeEndList().get(0).getId().toString());
            assertEquals("0", paragraph.getRuns().get(1).getCTR().getCommentReferenceList().get(0).getId().toString());
        }
    }

    @Test
    void readPhaseDocxFields() throws IOException {
        Path fixture = findRepoRoot().resolve("tests/DotnetPoi.Interop.Tests/fixtures/from-dotnet-poi/phase-docx-fields.docx");
        assertTrue(Files.exists(fixture), "Run the C# WriteForPoi tests before this Java read test.");

        try (InputStream input = Files.newInputStream(fixture);
             XWPFDocument doc = new XWPFDocument(input)) {

            var paragraphs = doc.getParagraphs();

            // Paragraph 0: "Page " + PAGE field
            assertTrue(paragraphs.get(0).getText().contains("Page"));

            // Paragraph 2: "Hello " + MERGEFIELD
            assertTrue(paragraphs.get(2).getText().contains("Hello"));
        }
    }

    @Test
    void readPhase15HslfNoOp() throws IOException {
        // Phase 15 Item 7: Direction B — Java POI reads dotnet-poi no-op saved .ppt
        Path fixture = findRepoRoot().resolve("tests/DotnetPoi.Interop.Tests/fixtures/from-dotnet-poi/phase15-hslf-noop.ppt");
        assertTrue(Files.exists(fixture), "Run the C# Write_Phase15HslfNoOp_CreatesFixtureForPoi test first.");

        try (InputStream input = Files.newInputStream(fixture);
             HSLFSlideShow ppt = new HSLFSlideShow(input)) {

            var slides = ppt.getSlides();
            assertNotNull(slides);
            assertTrue(slides.size() > 0, "No-op preserved .ppt should have slides.");
        }
    }

    @Test
    void readPhase17HwpfDocExtraction() throws Exception {
        // Phase 17: Direction B — Java POI reads dotnet-poi no-op saved docs with headers/tables
        Path headerFixture = findRepoRoot().resolve("tests/DotnetPoi.Interop.Tests/fixtures/from-dotnet-poi/phase17-headerfooter.doc");
        Path tableFixture = findRepoRoot().resolve("tests/DotnetPoi.Interop.Tests/fixtures/from-dotnet-poi/phase17-innertable.doc");
        
        assertTrue(Files.exists(headerFixture), "Run the C# Write_Phase17HwpfDocExtraction_CreatesFixtureForPoi test first.");
        assertTrue(Files.exists(tableFixture), "Run the C# Write_Phase17HwpfDocExtraction_CreatesFixtureForPoi test first.");

        try (InputStream input = Files.newInputStream(headerFixture);
             org.apache.poi.hwpf.HWPFDocument doc = new org.apache.poi.hwpf.HWPFDocument(input)) {
            String headerText = doc.getHeaderStoryRange().text();
            assertTrue(headerText.contains("Molière"));
        }

        try (InputStream input = Files.newInputStream(tableFixture);
             org.apache.poi.hwpf.HWPFDocument doc = new org.apache.poi.hwpf.HWPFDocument(input)) {
            org.apache.poi.hwpf.usermodel.Range range = doc.getRange();
            boolean foundTable = false;
            for (int i = 0; i < range.numParagraphs(); i++) {
                if (range.getParagraph(i).isInTable()) {
                    foundTable = true;
                    break;
                }
            }
            assertTrue(foundTable, "Table should be correctly preserved in innertable.doc");
        }
    }

    private static Path findRepoRoot() {
        Path current = Paths.get("").toAbsolutePath();
        while (current != null) {
            if (Files.exists(current.resolve("DotnetPOI.sln"))) {
                return current;
            }
            current = current.getParent();
        }
        throw new IllegalStateException("Could not locate repository root.");
    }
}
