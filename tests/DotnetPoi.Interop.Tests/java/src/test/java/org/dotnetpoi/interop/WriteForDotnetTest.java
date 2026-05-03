package org.dotnetpoi.interop;

import static org.junit.jupiter.api.Assertions.assertTrue;

import java.io.IOException;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;

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
