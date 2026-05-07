package org.dotnetpoi.interop;

import static org.junit.jupiter.api.Assertions.assertTrue;

import java.io.IOException;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.FormulaError;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.apache.poi.xwpf.usermodel.XWPFHeader;
import org.apache.poi.xwpf.usermodel.XWPFFooter;
import org.apache.poi.xwpf.usermodel.BreakType;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.sl.usermodel.*;
import org.apache.poi.xslf.usermodel.*;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.wp.usermodel.HeaderFooterType;
import org.junit.jupiter.api.Test;

public class WriteForDotnetTest {
    @Test
    void writePhase6BasicHssfWorkbook() throws IOException {
        Path fixture = findRepoRoot().resolve("tests/DotnetPoi.Interop.Tests/fixtures/from-poi/phase6-basic.xls");
        Files.createDirectories(fixture.getParent());

        try (HSSFWorkbook workbook = new HSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("From POI HSSF");
            Row row = sheet.createRow(0);
            row.createCell(0).setCellValue("from apache poi hssf");
            row.createCell(1).setCellValue(123.75);
            row.createCell(2).setCellValue(false);

            try (OutputStream output = Files.newOutputStream(fixture)) {
                workbook.write(output);
            }
        }

        assertTrue(Files.exists(fixture));
        assertTrue(Files.size(fixture) > 0);
    }

    @Test
    void writePhase13SampleDoc() throws IOException {
        // Phase 13: Direction A — Java POI writes SampleDoc.doc → dotnet-poi reads
        Path fixture = findRepoRoot().resolve("tests/DotnetPoi.Interop.Tests/fixtures/from-poi/phase13-sample-doc.doc");
        Files.createDirectories(fixture.getParent());

        try (org.apache.poi.hwpf.HWPFDocument doc = new org.apache.poi.hwpf.HWPFDocument()) {
            org.apache.poi.hwpf.usermodel.Range range = doc.getRange();
            range.insertBefore("I am a Java POI generated test document.\r");
            range.insertAfter("Second paragraph from Java POI.\r");

            try (OutputStream output = Files.newOutputStream(fixture)) {
                doc.write(output);
            }
        }

        assertTrue(Files.exists(fixture));
        assertTrue(Files.size(fixture) > 0);
    }

    @Test
    void writePhase12HssfStyles() throws IOException {
        // Phase 12 item 4: Direction A — Java POI writes .xls with styles
        Path fixture = findRepoRoot().resolve("tests/DotnetPoi.Interop.Tests/fixtures/from-poi/phase12-hssf-styles.xls");
        Files.createDirectories(fixture.getParent());

        try (HSSFWorkbook workbook = new HSSFWorkbook()) {
            Font boldFont = workbook.createFont();
            boldFont.setBold(true);
            boldFont.setFontName("Calibri");
            boldFont.setFontHeightInPoints((short) 14);
            boldFont.setItalic(true);

            CellStyle style1 = workbook.createCellStyle();
            style1.setFont(boldFont);
            style1.setAlignment(HorizontalAlignment.CENTER);
            style1.setWrapText(true);
            style1.setBorderBottom(BorderStyle.THIN);
            style1.setDataFormat(workbook.createDataFormat().getFormat("0.00"));

            CellStyle style2 = workbook.createCellStyle();
            style2.setAlignment(HorizontalAlignment.RIGHT);
            style2.setBorderLeft(BorderStyle.MEDIUM);

            Sheet sheet = workbook.createSheet("Styles");
            Row row = sheet.createRow(0);
            Cell cell0 = row.createCell(0);
            cell0.setCellValue(42.5);
            cell0.setCellStyle(style1);

            Cell cell1 = row.createCell(1);
            cell1.setCellValue("right");
            cell1.setCellStyle(style2);

            Cell cell2 = row.createCell(2);
            cell2.setCellValue("no style");

            try (OutputStream output = Files.newOutputStream(fixture)) {
                workbook.write(output);
            }
        }

        assertTrue(Files.exists(fixture));
        assertTrue(Files.size(fixture) > 0);
    }

    @Test
    void writePhase12HssfLayout() throws IOException {
        // Phase 12 item 4: Direction A — Java POI writes .xls with layout
        Path fixture = findRepoRoot().resolve("tests/DotnetPoi.Interop.Tests/fixtures/from-poi/phase12-hssf-layout.xls");
        Files.createDirectories(fixture.getParent());

        try (HSSFWorkbook workbook = new HSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Layout");
            // Column widths (in 1/256 of character width)
            sheet.setColumnWidth(0, 5000);
            sheet.setColumnWidth(1, 8000);
            // Hidden column
            sheet.setColumnHidden(2, true);
            // Row heights
            Row row0 = sheet.createRow(0);
            row0.setHeightInPoints(30);
            row0.createCell(0).setCellValue("wide col");
            Row row1 = sheet.createRow(1);
            row1.setZeroHeight(true);
            row1.createCell(0).setCellValue("hidden row");
            Row row2 = sheet.createRow(2);
            row2.setHeightInPoints(20);
            row2.createCell(0).setCellValue("normal");
            // Merged regions
            sheet.addMergedRegion(new org.apache.poi.ss.util.CellRangeAddress(3, 3, 0, 2));
            Row row3 = sheet.createRow(3);
            row3.createCell(0).setCellValue("merged");
            // Freeze pane
            sheet.createFreezePane(1, 2);

            try (OutputStream output = Files.newOutputStream(fixture)) {
                workbook.write(output);
            }
        }

        assertTrue(Files.exists(fixture));
        assertTrue(Files.size(fixture) > 0);
    }

    @Test
    void writePhase12HssfUnicode() throws IOException {
        // Phase 12 item 3: Unicode/Japanese sheet names and string cells
        Path fixture = findRepoRoot().resolve("tests/DotnetPoi.Interop.Tests/fixtures/from-poi/phase12-hssf-unicode.xls");
        Files.createDirectories(fixture.getParent());

        try (HSSFWorkbook workbook = new HSSFWorkbook()) {
            Sheet sheet1 = workbook.createSheet("日本語");
            Row row0 = sheet1.createRow(0);
            row0.createCell(0).setCellValue("テスト文字列");
            row0.createCell(1).setCellValue("hello 世界");
            row0.createCell(2).setCellValue("こんにちは");

            Sheet sheet2 = workbook.createSheet("中文测试");
            sheet2.createRow(0).createCell(0).setCellValue("汉字测试");

            try (OutputStream output = Files.newOutputStream(fixture)) {
                workbook.write(output);
            }
        }

        assertTrue(Files.exists(fixture));
        assertTrue(Files.size(fixture) > 0);
    }

    @Test
    void writePhase12HssfComprehensive() throws IOException {
        Path fixture = findRepoRoot().resolve("tests/DotnetPoi.Interop.Tests/fixtures/from-poi/phase12-hssf-comprehensive.xls");
        Files.createDirectories(fixture.getParent());

        try (HSSFWorkbook workbook = new HSSFWorkbook()) {
            // Sheet 1: all cell types in one row
            Sheet sheet1 = workbook.createSheet("CellTypes");
            Row row0 = sheet1.createRow(0);
            row0.createCell(0).setCellValue("string value");
            row0.createCell(1).setCellValue(42.5);
            row0.createCell(2).setCellValue(true);
            row0.createCell(3).setCellValue(false);
            row0.createCell(4).setCellErrorValue(FormulaError.DIV0.getCode());
            row0.createCell(5).setCellErrorValue(FormulaError.NA.getCode());
            row0.createCell(6); // blank cell

            // Sheet 2: sparse layout
            Sheet sheet2 = workbook.createSheet("Sparse");
            sheet2.createRow(0).createCell(0).setCellValue("row0col0");
            sheet2.createRow(5).createCell(3).setCellValue(99.9);
            sheet2.createRow(10).createCell(0).setCellValue("row10");

            try (OutputStream output = Files.newOutputStream(fixture)) {
                workbook.write(output);
            }
        }

        assertTrue(Files.exists(fixture));
        assertTrue(Files.size(fixture) > 0);
    }

    @Test
    void writePhase1BasicWorkbook() throws IOException {
        Path fixture = findRepoRoot().resolve("tests/DotnetPoi.Interop.Tests/fixtures/from-poi/phase1-basic.xlsx");
        Files.createDirectories(fixture.getParent());

        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            Sheet data = workbook.createSheet("From POI");
            Row row = data.createRow(0);
            row.createCell(0).setCellValue("from apache poi");
            row.createCell(1).setCellValue(123.25);
            row.createCell(2).setCellValue(0.0);
            data.createRow(1).createCell(0).setCellValue("second row");

            Sheet second = workbook.createSheet("Second");
            second.createRow(2).createCell(3).setCellValue(99.0);

            try (OutputStream output = Files.newOutputStream(fixture)) {
                workbook.write(output);
            }
        }

        assertTrue(Files.exists(fixture));
        assertTrue(Files.size(fixture) > 0);
    }

    @Test
    void writePhase7FormulaWorkbook() throws IOException {
        Path fixture = findRepoRoot().resolve("tests/DotnetPoi.Interop.Tests/fixtures/from-poi/phase7-formulas.xlsx");
        Files.createDirectories(fixture.getParent());

        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Formulas");

            // Row 0: base values
            Row r0 = sheet.createRow(0);
            r0.createCell(0).setCellValue(10.0);
            r0.createCell(1).setCellValue(20.0);

            // Row 1: numeric formula (SUM)
            sheet.createRow(1).createCell(0).setCellFormula("A1+B1");

            // Row 2: string formula (CONCATENATE) — t="str"
            sheet.createRow(2).createCell(0).setCellFormula("\"hello \"&\"world\"");

            // Row 3: boolean cell (not formula)
            sheet.createRow(3).createCell(0).setCellValue(true);

            // Row 4: error formula
            sheet.createRow(4).createCell(0).setCellFormula("1/0");

            // Evaluate formulas to populate cached <v> values
            FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
            for (Row row : sheet) {
                for (Cell cell : row) {
                    if (cell.getCellType() == org.apache.poi.ss.usermodel.CellType.FORMULA) {
                        evaluator.evaluateFormulaCell(cell);
                    }
                }
            }

            try (OutputStream output = Files.newOutputStream(fixture)) {
                workbook.write(output);
            }
        }

        assertTrue(Files.exists(fixture));
        assertTrue(Files.size(fixture) > 0);
    }

    @Test
    void writePhase5Step2RecalcWorkbook() throws IOException {
        Path fixture = findRepoRoot().resolve("tests/DotnetPoi.Interop.Tests/fixtures/from-poi/phase5-step2-recalc.xlsx");
        Files.createDirectories(fixture.getParent());

        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Recalc");
            sheet.createRow(0).createCell(0).setCellFormula("B1+C1");
            workbook.setForceFormulaRecalculation(true);

            try (OutputStream output = Files.newOutputStream(fixture)) {
                workbook.write(output);
            }
        }

        assertTrue(Files.exists(fixture));
        assertTrue(Files.size(fixture) > 0);
    }

    @Test
    void writePhaseDocxComprehensive() throws IOException {
        Path fixture = findRepoRoot().resolve("tests/DotnetPoi.Interop.Tests/fixtures/from-poi/phase-docx-comprehensive.docx");
        Files.createDirectories(fixture.getParent());

        try (XWPFDocument doc = new XWPFDocument()) {
            // --- Header ---
            XWPFHeader header = doc.createHeader(HeaderFooterType.DEFAULT);
            XWPFParagraph headerPara = header.createParagraph();
            headerPara.setAlignment(ParagraphAlignment.CENTER);
            XWPFRun headerRun = headerPara.createRun();
            headerRun.setText("Interop Header");

            // --- Paragraph 1: plain text ---
            XWPFParagraph p1 = doc.createParagraph();
            p1.setAlignment(ParagraphAlignment.LEFT);
            XWPFRun r1 = p1.createRun();
            r1.setText("First paragraph");

            // --- Paragraph 2: bold + normal ---
            XWPFParagraph p2 = doc.createParagraph();
            XWPFRun r2b = p2.createRun();
            r2b.setBold(true);
            r2b.setText("Bold");
            XWPFRun r2n = p2.createRun();
            r2n.setText(" and normal");

            // --- Paragraph 3: italic ---
            XWPFParagraph p3 = doc.createParagraph();
            XWPFRun r3 = p3.createRun();
            r3.setItalic(true);
            r3.setText("Italic text");

            // --- Table (2x2) ---
            XWPFTable table = doc.createTable(2, 2);
            table.getRow(0).getCell(0).setText("A1");
            table.getRow(0).getCell(1).setText("B1");
            table.getRow(1).getCell(0).setText("A2");
            table.getRow(1).getCell(1).setText("B2");

            // --- Paragraph 4: hyperlink ---
            XWPFParagraph linkPara = doc.createParagraph();
            XWPFRun linkRun = linkPara.createRun();
            linkRun.setText("Click here for Apache POI");

            // --- Footer ---
            XWPFFooter footer = doc.createFooter(HeaderFooterType.DEFAULT);
            XWPFParagraph footerPara = footer.createParagraph();
            footerPara.setAlignment(ParagraphAlignment.CENTER);
            XWPFRun footerRun = footerPara.createRun();
            footerRun.setText("Interop Footer");

            try (java.io.OutputStream output = Files.newOutputStream(fixture)) {
                doc.write(output);
            }
        }

        assertTrue(Files.exists(fixture));
        assertTrue(Files.size(fixture) > 0);
    }

    @Test
    void writePhasePptxComprehensive() throws IOException {
        Path fixture = findRepoRoot().resolve("tests/DotnetPoi.Interop.Tests/fixtures/from-poi/phase-pptx-comprehensive.pptx");
        Files.createDirectories(fixture.getParent());

        try (XMLSlideShow prs = new XMLSlideShow()) {
            XSLFSlide slide = prs.createSlide();

            // --- Text box with formatted text ---
            XSLFAutoShape tb = slide.createTextBox();
            tb.setAnchor(new java.awt.Rectangle(100000, 100000, 4000000, 500000));

            XSLFTextParagraph p1 = tb.addNewTextParagraph();
            XSLFTextRun r1 = p1.addNewTextRun();
            r1.setText("Bold Title");
            r1.setBold(true);
            r1.setFontSize(18.0);

            XSLFTextParagraph p2 = tb.addNewTextParagraph();
            XSLFTextRun r2 = p2.addNewTextRun();
            r2.setText("Italic subtitle");
            r2.setItalic(true);
            r2.setFontSize(14.0);

            try (java.io.OutputStream output = Files.newOutputStream(fixture)) {
                prs.write(output);
            }
        }

        assertTrue(Files.exists(fixture));
        assertTrue(Files.size(fixture) > 0);
    }

    @Test
    void writePhaseAutoFilterSheet() throws IOException {
        Path fixture = findRepoRoot().resolve("tests/DotnetPoi.Interop.Tests/fixtures/from-poi/phase-autofilter.xlsx");
        Files.createDirectories(fixture.getParent());

        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Data");
            Row header = sheet.createRow(0);
            header.createCell(0).setCellValue("Category");
            header.createCell(1).setCellValue("Value");

            Row food = sheet.createRow(1);
            food.createCell(0).setCellValue("Food");
            food.createCell(1).setCellValue(100);

            Row travel = sheet.createRow(2);
            travel.createCell(0).setCellValue("Travel");
            travel.createCell(1).setCellValue(200);
            sheet.setAutoFilter(new org.apache.poi.ss.util.CellRangeAddress(0, 2, 0, 1));

            try (OutputStream output = Files.newOutputStream(fixture)) {
                workbook.write(output);
            }
        }

        assertTrue(Files.exists(fixture));
        assertTrue(Files.size(fixture) > 0);
    }

    @Test
    void writePhaseProtectedSheet() throws IOException {
        Path fixture = findRepoRoot().resolve("tests/DotnetPoi.Interop.Tests/fixtures/from-poi/phase-protection.xlsx");
        Files.createDirectories(fixture.getParent());

        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Data");
            sheet.createRow(0).createCell(0).setCellValue("protected cell");
            sheet.protectSheet("password");
            workbook.lockStructure();

            try (OutputStream output = Files.newOutputStream(fixture)) {
                workbook.write(output);
            }
        }

        assertTrue(Files.exists(fixture));
        assertTrue(Files.size(fixture) > 0);
    }

    @Test
    void writePhaseActiveSheet() throws IOException {
        Path fixture = findRepoRoot().resolve("tests/DotnetPoi.Interop.Tests/fixtures/from-poi/phase-active-sheet.xlsx");
        Files.createDirectories(fixture.getParent());

        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            workbook.createSheet("First");
            workbook.createSheet("Second");
            workbook.createSheet("Third");
            workbook.setActiveSheet(1);

            try (OutputStream output = Files.newOutputStream(fixture)) {
                workbook.write(output);
            }
        }

        assertTrue(Files.exists(fixture));
        assertTrue(Files.size(fixture) > 0);
    }

    @Test
    void writePhaseDocxWithFields() throws IOException {
        // Skip: POI 5.5.1 does not support addNewFld() / XWPF Field API.
        // The Direction A C# test (Read_DocxWithFields_GeneratedByPoi) is already Skipped.
        // Keep this method body minimal so the test compiles and passes gracefully.
        Path fixture = findRepoRoot().resolve("tests/DotnetPoi.Interop.Tests/fixtures/from-poi/phase-docx-fields.docx");
        Files.createDirectories(fixture.getParent());

        try (XWPFDocument doc = new XWPFDocument()) {
            // Create a simple docx with text (no field codes - not supported in POI 5.5.1)
            doc.createParagraph().createRun().setText("Page placeholder");
            doc.createParagraph().createRun().setText("TOC placeholder");
            doc.createParagraph().createRun().setText("Hello MERGEFIELD placeholder");

            try (OutputStream output = Files.newOutputStream(fixture)) {
                doc.write(output);
            }
        }

        assertTrue(Files.exists(fixture));
        assertTrue(Files.size(fixture) > 0);
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
