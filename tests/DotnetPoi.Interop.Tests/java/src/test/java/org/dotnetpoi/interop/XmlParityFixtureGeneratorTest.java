package org.dotnetpoi.interop;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.junit.jupiter.api.Test;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.file.DirectoryStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.StandardCopyOption;
import java.util.Base64;
import java.util.Enumeration;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;

public class XmlParityFixtureGeneratorTest {
    private static final byte[] ONE_BY_ONE_PNG = Base64.getDecoder().decode(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO2O8WcAAAAASUVORK5CYII="
    );

    @Test
    public void GenerateFixtures() throws Exception {
        Path fixturesRoot = getFixturesRoot();
        Path workbooksRoot = fixturesRoot.resolve("_workbooks");
        Files.createDirectories(workbooksRoot);
        Files.createDirectories(fixturesRoot);

        generateCellZeroCases(fixturesRoot, workbooksRoot);
        generateEmptyCellCases(fixturesRoot, workbooksRoot);
        generateCellTypeCases(fixturesRoot, workbooksRoot);
        generateMultiSheetCases(fixturesRoot, workbooksRoot);
        generateNamespaceCases(fixturesRoot, workbooksRoot);
        generateInlineStringCases(fixturesRoot, workbooksRoot);
        generateDocxBasicCases(fixturesRoot, workbooksRoot);
        generatePptxBasicCases(fixturesRoot, workbooksRoot);
        generateMacroEnabledCases(fixturesRoot, workbooksRoot);
    }

    private static Path getFixturesRoot() {
        Path directory = Paths.get("").toAbsolutePath();
        while (directory != null) {
            if (Files.exists(directory.resolve("DotnetPOI.sln"))) {
                return directory.resolve("tests")
                        .resolve("DotnetPoi.Interop.Tests")
                        .resolve("fixtures")
                        .resolve("xml-parity");
            }
            directory = directory.getParent();
        }
        throw new IllegalStateException("Could not locate the dotnet-poi repository root.");
    }

    private static Path getRepoRoot() {
        Path directory = Paths.get("").toAbsolutePath();
        while (directory != null) {
            if (Files.exists(directory.resolve("DotnetPOI.sln"))) {
                return directory;
            }
            directory = directory.getParent();
        }
        throw new IllegalStateException("Could not locate the dotnet-poi repository root.");
    }

    private static void generateCellZeroCases(Path fixturesRoot, Path workbooksRoot) throws IOException {
        String caseName = "cell-zero";
        cleanupCaseFiles(fixturesRoot, caseName);

        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Zeroes");
            Row row = sheet.createRow(0);

            row.createCell(0).setCellValue(0);
            row.createCell(1).setCellValue(0.0d);
            row.createCell(2).setCellValue(-0.0d);
            row.createCell(3).setCellValue(0.0000d);

            writeAndExtract(caseName, workbook, fixturesRoot, workbooksRoot);
        }
    }

    private static void generateEmptyCellCases(Path fixturesRoot, Path workbooksRoot) throws IOException {
        String caseName = "cell-empty";
        cleanupCaseFiles(fixturesRoot, caseName);

        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Empty");
            Row row = sheet.createRow(0);

            Cell blankCell = row.createCell(0);
            blankCell.setBlank();
            row.createCell(1).setCellValue("");
            sheet.createRow(1); // empty row with no cells

            writeAndExtract(caseName, workbook, fixturesRoot, workbooksRoot);
        }
    }

    private static void generateCellTypeCases(Path fixturesRoot, Path workbooksRoot) throws IOException {
        String caseName = "cell-types";
        cleanupCaseFiles(fixturesRoot, caseName);

        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            CreationHelper helper = workbook.getCreationHelper();
            Sheet sheet = workbook.createSheet("Types");
            Row row = sheet.createRow(0);

            row.createCell(0).setCellValue("Text");
            row.createCell(1).setCellValue(42.5d);
            row.createCell(2).setCellValue(true);

            Cell dateCell = row.createCell(3);
            // Set the date as an Excel serial number (double) to avoid timezone-dependent
            // Date→serial conversion. 25569.0 = 1970-01-01 midnight in Excel date format.
            // Using new Date(0) would produce 25569.375 in JST (UTC+9) because POI converts
            // via the JVM's default timezone, making the fixture non-reproducible across
            // timezones.
            dateCell.setCellValue(25569.0);
            CellStyle dateStyle = workbook.createCellStyle();
            dateStyle.setDataFormat(helper.createDataFormat().getFormat("yyyy-mm-dd"));
            dateCell.setCellStyle(dateStyle);

            Cell formulaCell = row.createCell(4);
            formulaCell.setCellFormula("SUM(1,2,3)");

            writeAndExtract(caseName, workbook, fixturesRoot, workbooksRoot);
        }
    }

    private static void generateMultiSheetCases(Path fixturesRoot, Path workbooksRoot) throws IOException {
        String caseName = "multi-sheet";
        cleanupCaseFiles(fixturesRoot, caseName);

        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            Sheet sheet1 = workbook.createSheet("Alpha");
            Sheet sheet2 = workbook.createSheet("Beta");
            Sheet sheet3 = workbook.createSheet("Gamma");

            sheet1.createRow(0).createCell(0).setCellValue("A1");
            sheet2.createRow(1).createCell(1).setCellValue(123);
            sheet3.createRow(2).createCell(2).setCellValue(false);

            writeAndExtract(caseName, workbook, fixturesRoot, workbooksRoot);
        }
    }

    private static void generateNamespaceCases(Path fixturesRoot, Path workbooksRoot) throws IOException {
        String caseName = "namespaces";
        cleanupCaseFiles(fixturesRoot, caseName);

        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            CreationHelper helper = workbook.getCreationHelper();
            Sheet sheet = workbook.createSheet("Namespaces");
            Row row = sheet.createRow(0);
            Cell cell = row.createCell(0);
            cell.setCellValue("Image + Comment");

            Drawing<?> drawing = sheet.createDrawingPatriarch();
            ClientAnchor anchor = helper.createClientAnchor();
            anchor.setCol1(0);
            anchor.setRow1(0);

            int pictureId = workbook.addPicture(ONE_BY_ONE_PNG, Workbook.PICTURE_TYPE_PNG);
            drawing.createPicture(anchor, pictureId);

            Comment comment = drawing.createCellComment(anchor);
            RichTextString richText = helper.createRichTextString("Comment text");
            comment.setString(richText);
            comment.setAuthor("poi");
            cell.setCellComment(comment);

            writeAndExtract(caseName, workbook, fixturesRoot, workbooksRoot);
        }
    }

    private static void generateInlineStringCases(Path fixturesRoot, Path workbooksRoot) throws IOException {
        String caseName = "inline-strings";
        cleanupCaseFiles(fixturesRoot, caseName);

        Path workbookPath = workbooksRoot.resolve(caseName + ".xlsx");
        if (!Files.exists(workbookPath)) {
            throw new IOException("Missing inline string workbook fixture: " + workbookPath);
        }
        extractXmlEntries(workbookPath, caseName, fixturesRoot);
    }

    private static void generateDocxBasicCases(Path fixturesRoot, Path workbooksRoot) throws IOException {
        String caseName = "docx-basic";
        cleanupCaseFiles(fixturesRoot, caseName);

        try (XWPFDocument doc = new XWPFDocument()) {
            XWPFParagraph para = doc.createParagraph();
            para.createRun().setText("docx parity");

            Path outputPath = workbooksRoot.resolve(caseName + ".docx");
            writeAndExtractPackage(caseName, outputPath, doc::write, fixturesRoot);
        }
    }

    private static void generatePptxBasicCases(Path fixturesRoot, Path workbooksRoot) throws IOException {
        String caseName = "pptx-basic";
        cleanupCaseFiles(fixturesRoot, caseName);

        try (XMLSlideShow show = new XMLSlideShow()) {
            show.createSlide();

            Path outputPath = workbooksRoot.resolve(caseName + ".pptx");
            writeAndExtractPackage(caseName, outputPath, show::write, fixturesRoot);
        }
    }

    private static void generateMacroEnabledCases(Path fixturesRoot, Path workbooksRoot) throws IOException {
        Path repoRoot = getRepoRoot();

        generateMacroEnabledXlsm(fixturesRoot, workbooksRoot,
                repoRoot.resolve("tests/test-files/example.xlsm"));
        generateMacroEnabledDocm(fixturesRoot, workbooksRoot,
                repoRoot.resolve("tests/test-files/example.docm"));
        generateMacroEnabledPptm(fixturesRoot, workbooksRoot,
                repoRoot.resolve("tests/test-files/example.pptm"));
    }

    private static void generateMacroEnabledXlsm(Path fixturesRoot, Path workbooksRoot, Path source) throws IOException {
        String caseName = "xlsm-basic";
        cleanupCaseFiles(fixturesRoot, caseName);

        try (XSSFWorkbook workbook = new XSSFWorkbook(Files.newInputStream(source))) {
            Path outputPath = workbooksRoot.resolve(caseName + ".xlsm");
            writeAndExtractPackage(caseName, outputPath, workbook::write, fixturesRoot);
        }
    }

    private static void generateMacroEnabledDocm(Path fixturesRoot, Path workbooksRoot, Path source) throws IOException {
        String caseName = "docm-basic";
        cleanupCaseFiles(fixturesRoot, caseName);

        try (XWPFDocument doc = new XWPFDocument(Files.newInputStream(source))) {
            Path outputPath = workbooksRoot.resolve(caseName + ".docm");
            writeAndExtractPackage(caseName, outputPath, doc::write, fixturesRoot);
        }
    }

    private static void generateMacroEnabledPptm(Path fixturesRoot, Path workbooksRoot, Path source) throws IOException {
        String caseName = "pptm-basic";
        cleanupCaseFiles(fixturesRoot, caseName);

        try (XMLSlideShow show = new XMLSlideShow(Files.newInputStream(source))) {
            Path outputPath = workbooksRoot.resolve(caseName + ".pptm");
            writeAndExtractPackage(caseName, outputPath, show::write, fixturesRoot);
        }
    }

    private static void cleanupCaseFiles(Path fixturesRoot, String caseName) throws IOException {
        if (!Files.exists(fixturesRoot)) {
            return;
        }
        try (DirectoryStream<Path> stream = Files.newDirectoryStream(fixturesRoot, caseName + "__*")) {
            for (Path path : stream) {
                Files.deleteIfExists(path);
            }
        }
    }

    private static void writeAndExtract(String caseName,
                                        Workbook workbook,
                                        Path fixturesRoot,
                                        Path workbooksRoot) throws IOException {
        Path workbookPath = workbooksRoot.resolve(caseName + ".xlsx");
        try (OutputStream outputStream = Files.newOutputStream(workbookPath)) {
            workbook.write(outputStream);
        }

        extractXmlEntries(workbookPath, caseName, fixturesRoot);
    }

    @FunctionalInterface
    private interface PackageWriter {
        void write(OutputStream outputStream) throws IOException;
    }

    private static void writeAndExtractPackage(String caseName,
                                               Path outputPath,
                                               PackageWriter writer,
                                               Path fixturesRoot) throws IOException {
        try (OutputStream outputStream = Files.newOutputStream(outputPath)) {
            writer.write(outputStream);
        }

        extractXmlEntries(outputPath, caseName, fixturesRoot);
    }

    private static void extractXmlEntries(Path workbookPath, String caseName, Path fixturesRoot) throws IOException {
        try (ZipFile zipFile = new ZipFile(workbookPath.toFile())) {
            Enumeration<? extends ZipEntry> entries = zipFile.entries();
            while (entries.hasMoreElements()) {
                ZipEntry entry = entries.nextElement();
                String name = entry.getName();
                if (!name.endsWith(".xml") && !name.endsWith(".rels")) {
                    continue;
                }

                String flatName = caseName + "__" + name.replace("/", "__");
                Path outputPath = fixturesRoot.resolve(flatName);
                try (InputStream inputStream = zipFile.getInputStream(entry)) {
                    Files.copy(inputStream, outputPath, StandardCopyOption.REPLACE_EXISTING);
                }
            }
        }
    }
}
