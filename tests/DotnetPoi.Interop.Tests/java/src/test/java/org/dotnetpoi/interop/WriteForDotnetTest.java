package org.dotnetpoi.interop;

import static org.junit.jupiter.api.Assertions.assertTrue;

import java.io.IOException;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
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
