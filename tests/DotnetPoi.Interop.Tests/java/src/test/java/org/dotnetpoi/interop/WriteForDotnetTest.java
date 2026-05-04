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
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

public class WriteForDotnetTest {
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
