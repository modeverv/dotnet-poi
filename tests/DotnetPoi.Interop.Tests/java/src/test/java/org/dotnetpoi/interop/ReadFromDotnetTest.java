package org.dotnetpoi.interop;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

public class ReadFromDotnetTest {
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
