package org.dotnetpoi.interop;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.PictureData;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.poifs.crypt.Decryptor;
import org.apache.poi.poifs.crypt.EncryptionInfo;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFPictureShape;
import org.apache.poi.xslf.usermodel.XSLFShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFPicture;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.junit.jupiter.api.Test;

public class ReadFromDotnetTest {
    private static byte[] loadTestImage() throws IOException {
        return Files.readAllBytes(findRepoRoot().resolve("tests/image.jpg"));
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
