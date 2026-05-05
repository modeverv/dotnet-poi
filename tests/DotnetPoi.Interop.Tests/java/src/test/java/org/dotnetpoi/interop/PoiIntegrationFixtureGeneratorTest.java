package org.dotnetpoi.interop;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFCreationHelper;
import org.apache.poi.xssf.usermodel.XSSFHyperlink;
import org.apache.poi.xssf.usermodel.XSSFName;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.DirectoryStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.StandardCopyOption;
import java.util.Enumeration;
import java.util.List;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;

public class PoiIntegrationFixtureGeneratorTest {
    @Test
    public void GenerateSharedStringsBasicFixture() throws Exception {
        Path fixturesRoot = getFixturesRoot();
        Path workbooksRoot = fixturesRoot.resolve("_workbooks");
        Files.createDirectories(fixturesRoot);
        Files.createDirectories(workbooksRoot);

        String caseName = "poi-integration-shared-strings-basic";
        cleanupCaseFiles(fixturesRoot, caseName);

        Path sourceWorkbook = getRepoRoot()
                .resolve("poi")
                .resolve("test-data")
                .resolve("spreadsheet")
                .resolve("sample.xlsx");

        Path outputWorkbook = workbooksRoot.resolve(caseName + ".xlsx");
        try (XSSFWorkbook workbook = new XSSFWorkbook(Files.newInputStream(sourceWorkbook));
             OutputStream outputStream = Files.newOutputStream(outputWorkbook)) {
            workbook.write(outputStream);
        }

        extractXmlEntries(outputWorkbook, caseName, fixturesRoot);
    }

    @Test
    public void GenerateSharedStringsEscapingFixture() throws Exception {
        Path fixturesRoot = getFixturesRoot();
        Path workbooksRoot = fixturesRoot.resolve("_workbooks");
        Files.createDirectories(fixturesRoot);
        Files.createDirectories(workbooksRoot);

        String caseName = "poi-integration-shared-strings-escaping";
        cleanupCaseFiles(fixturesRoot, caseName);

        Path sourceText = getRepoRoot()
                .resolve("poi")
                .resolve("test-data")
                .resolve("spreadsheet")
                .resolve("48936-strings.txt");

        Path outputWorkbook = workbooksRoot.resolve(caseName + ".xlsx");
        try (XSSFWorkbook workbook = new XSSFWorkbook();
             OutputStream outputStream = Files.newOutputStream(outputWorkbook)) {
            setStableCoreProperties(workbook);

            Sheet sheet = workbook.createSheet("strings");
            List<String> lines = Files.readAllLines(sourceText, StandardCharsets.UTF_8);
            int rowIndex = 0;
            for (String line : lines) {
                String value = line.trim();
                if (!value.isEmpty()) {
                    sheet.createRow(rowIndex++).createCell(0).setCellValue(value);
                }
            }

            workbook.write(outputStream);
        }

        extractXmlEntries(outputWorkbook, caseName, fixturesRoot);
    }

    @Test
    public void GenerateStylesFormattingFixture() throws Exception {
        Path fixturesRoot = getFixturesRoot();
        Path workbooksRoot = fixturesRoot.resolve("_workbooks");
        Files.createDirectories(fixturesRoot);
        Files.createDirectories(workbooksRoot);

        String caseName = "poi-integration-styles-formatting";
        cleanupCaseFiles(fixturesRoot, caseName);

        Path sourceWorkbook = getRepoRoot()
                .resolve("poi")
                .resolve("test-data")
                .resolve("spreadsheet")
                .resolve("Formatting.xlsx");

        Path outputWorkbook = workbooksRoot.resolve(caseName + ".xlsx");
        try (XSSFWorkbook workbook = new XSSFWorkbook(Files.newInputStream(sourceWorkbook));
             OutputStream outputStream = Files.newOutputStream(outputWorkbook)) {
            workbook.write(outputStream);
        }

        extractXmlEntries(outputWorkbook, caseName, fixturesRoot);
    }

    @Test
    public void GenerateCommentsWriteReadFixture() throws Exception {
        Path fixturesRoot = getFixturesRoot();
        Path workbooksRoot = fixturesRoot.resolve("_workbooks");
        Files.createDirectories(fixturesRoot);
        Files.createDirectories(workbooksRoot);

        String caseName = "poi-integration-comments-write-read";
        cleanupCaseFiles(fixturesRoot, caseName);

        Path sourceWorkbook = getRepoRoot()
                .resolve("poi")
                .resolve("test-data")
                .resolve("spreadsheet")
                .resolve("WithVariousData.xlsx");

        Path outputWorkbook = workbooksRoot.resolve(caseName + ".xlsx");
        try (XSSFWorkbook workbook = new XSSFWorkbook(Files.newInputStream(sourceWorkbook));
             OutputStream outputStream = Files.newOutputStream(outputWorkbook)) {
            Sheet sheet1 = workbook.getSheetAt(0);
            Sheet sheet2 = workbook.getSheetAt(1);

            Row row5 = sheet1.getRow(4);
            Comment existingComment = row5.getCell(2).getCellComment();
            existingComment.setAuthor("Apache POI");
            existingComment.setString(new XSSFRichTextString("Hello!"));

            Row row2Sheet2 = sheet2.createRow(2);
            Cell cellB3 = row2Sheet2.createCell(1);
            Drawing<?> drawing = sheet2.createDrawingPatriarch();
            Comment newComment = drawing.createCellComment(new XSSFClientAnchor());
            newComment.setAuthor("Also POI");
            newComment.setString(new XSSFRichTextString("A new comment"));
            cellB3.setCellComment(newComment);

            workbook.write(outputStream);
        }

        extractXmlEntries(outputWorkbook, caseName, fixturesRoot);
    }

    @Test
    public void GeneratePicturesMultiSheetFixture() throws Exception {
        Path fixturesRoot = getFixturesRoot();
        Path workbooksRoot = fixturesRoot.resolve("_workbooks");
        Files.createDirectories(fixturesRoot);
        Files.createDirectories(workbooksRoot);

        String caseName = "poi-integration-pictures-multi-sheet";
        cleanupCaseFiles(fixturesRoot, caseName);

        Path outputWorkbook = workbooksRoot.resolve(caseName + ".xlsx");
        try (XSSFWorkbook workbook = new XSSFWorkbook();
             OutputStream outputStream = Files.newOutputStream(outputWorkbook)) {
            setStableCoreProperties(workbook);

            Sheet sheet1 = workbook.createSheet("Sheet 1");
            Sheet sheet2 = workbook.createSheet("Sheet 2");
            byte[] pic1Data = "picture1".getBytes(StandardCharsets.ISO_8859_1);
            byte[] pic2Data = "picture2".getBytes(StandardCharsets.ISO_8859_1);
            int pic1 = workbook.addPicture(pic1Data, XSSFWorkbook.PICTURE_TYPE_JPEG);
            int pic2 = workbook.addPicture(pic2Data, XSSFWorkbook.PICTURE_TYPE_PNG);

            Drawing<?> drawing1 = sheet1.createDrawingPatriarch();
            drawing1.createPicture(new XSSFClientAnchor(), pic1);
            drawing1.createPicture(new XSSFClientAnchor(), pic2);

            Drawing<?> drawing2 = sheet2.createDrawingPatriarch();
            drawing2.createPicture(new XSSFClientAnchor(), pic2);
            drawing2.createPicture(new XSSFClientAnchor(), pic1);

            workbook.write(outputStream);
        }

        extractXmlEntries(outputWorkbook, caseName, fixturesRoot);
    }

    @Test
    public void GenerateRelationshipsHyperlinksCommentsFixture() throws Exception {
        Path fixturesRoot = getFixturesRoot();
        Path workbooksRoot = fixturesRoot.resolve("_workbooks");
        Files.createDirectories(fixturesRoot);
        Files.createDirectories(workbooksRoot);

        String caseName = "poi-integration-relationships-hyperlinks-comments";
        cleanupCaseFiles(fixturesRoot, caseName);

        Path sourceWorkbook = getRepoRoot()
                .resolve("poi")
                .resolve("test-data")
                .resolve("openxml4j")
                .resolve("ExcelWithHyperlinks.xlsx");

        Path outputWorkbook = workbooksRoot.resolve(caseName + ".xlsx");
        try (XSSFWorkbook workbook = new XSSFWorkbook(Files.newInputStream(sourceWorkbook));
             OutputStream outputStream = Files.newOutputStream(outputWorkbook)) {
            workbook.write(outputStream);
        }

        extractXmlEntries(outputWorkbook, caseName, fixturesRoot);
    }

    @Test
    public void GenerateXlsmVbaPreserveFixture() throws Exception {
        Path fixturesRoot = getFixturesRoot();
        Path workbooksRoot = fixturesRoot.resolve("_workbooks");
        Files.createDirectories(fixturesRoot);
        Files.createDirectories(workbooksRoot);

        String caseName = "poi-integration-xlsm-vba-preserve";
        cleanupCaseFiles(fixturesRoot, caseName);

        Path sourceWorkbook = getRepoRoot()
                .resolve("poi")
                .resolve("test-data")
                .resolve("spreadsheet")
                .resolve("45431.xlsm");

        Path outputWorkbook = workbooksRoot.resolve(caseName + ".xlsm");
        try (XSSFWorkbook workbook = new XSSFWorkbook(Files.newInputStream(sourceWorkbook));
             OutputStream outputStream = Files.newOutputStream(outputWorkbook)) {
            workbook.write(outputStream);
        }

        extractXmlEntries(outputWorkbook, caseName, fixturesRoot);
    }

    @Test
    public void GenerateRichTextSpacePreserveFixture() throws Exception {
        Path fixturesRoot = getFixturesRoot();
        Path workbooksRoot = fixturesRoot.resolve("_workbooks");
        Files.createDirectories(fixturesRoot);
        Files.createDirectories(workbooksRoot);

        String caseName = "poi-integration-rich-text-space-preserve";
        cleanupCaseFiles(fixturesRoot, caseName);

        Path outputWorkbook = workbooksRoot.resolve(caseName + ".xlsx");
        try (XSSFWorkbook workbook = new XSSFWorkbook();
             OutputStream outputStream = Files.newOutputStream(outputWorkbook)) {
            setStableCoreProperties(workbook);

            Sheet sheet = workbook.createSheet("rich text");
            Font font = workbook.createFont();
            font.setBold(true);

            XSSFRichTextString leadingTrailing = new XSSFRichTextString("  Apache");
            leadingTrailing.append(" POI");
            leadingTrailing.append(" ");
            sheet.createRow(0).createCell(0).setCellValue(leadingTrailing);

            XSSFRichTextString newline = new XSSFRichTextString("Incorrect\n Line-Breaking");
            newline.applyFont(0, 9, font);
            sheet.createRow(1).createCell(0).setCellValue(newline);

            XSSFRichTextString trailingNewline = new XSSFRichTextString("Tab\tseparated\n");
            trailingNewline.applyFont(0, 3, font);
            sheet.createRow(2).createCell(0).setCellValue(trailingNewline);

            XSSFRichTextString multipleNewlines = new XSSFRichTextString("\n\n\nNew Line\n\n");
            multipleNewlines.applyFont(0, 3, font);
            multipleNewlines.applyFont(11, 13, font);
            sheet.createRow(3).createCell(0).setCellValue(multipleNewlines);

            workbook.write(outputStream);
        }

        extractXmlEntries(outputWorkbook, caseName, fixturesRoot);
    }

    @Test
    public void GenerateDefinedNamesPrintTitlesFixture() throws Exception {
        Path fixturesRoot = getFixturesRoot();
        Path workbooksRoot = fixturesRoot.resolve("_workbooks");
        Files.createDirectories(fixturesRoot);
        Files.createDirectories(workbooksRoot);

        String caseName = "poi-integration-defined-names-print-titles";
        cleanupCaseFiles(fixturesRoot, caseName);

        Path outputWorkbook = workbooksRoot.resolve(caseName + ".xlsx");
        try (XSSFWorkbook workbook = new XSSFWorkbook();
             OutputStream outputStream = Files.newOutputStream(outputWorkbook)) {
            setStableCoreProperties(workbook);

            Sheet first = workbook.createSheet("First Sheet");
            first.createRow(0).createCell(0).setCellValue("named");
            first.setRepeatingRows(CellRangeAddress.valueOf("1:4"));
            first.setRepeatingColumns(CellRangeAddress.valueOf("A:A"));

            Sheet second = workbook.createSheet("SecondSheet");
            second.createRow(0).createCell(0).setCellValue("second");
            second.setRepeatingRows(CellRangeAddress.valueOf("1:1"));
            second.setRepeatingColumns(CellRangeAddress.valueOf("B:C"));

            XSSFName sheet1Name = workbook.createName();
            sheet1Name.setNameName("Sheet1A1");
            sheet1Name.setRefersToFormula("'First Sheet'!$A$1");

            workbook.write(outputStream);
        }

        extractXmlEntries(outputWorkbook, caseName, fixturesRoot);
    }

    @Test
    public void GenerateHyperlinksFixture() throws Exception {
        Path fixturesRoot = getFixturesRoot();
        Path workbooksRoot = fixturesRoot.resolve("_workbooks");
        Files.createDirectories(fixturesRoot);
        Files.createDirectories(workbooksRoot);

        String caseName = "poi-integration-hyperlinks";
        cleanupCaseFiles(fixturesRoot, caseName);

        Path outputWorkbook = workbooksRoot.resolve(caseName + ".xlsx");
        try (XSSFWorkbook workbook = new XSSFWorkbook();
             OutputStream outputStream = Files.newOutputStream(outputWorkbook)) {
            setStableCoreProperties(workbook);

            Sheet sheet = workbook.createSheet("links");
            Row row = sheet.createRow(0);
            XSSFCreationHelper helper = workbook.getCreationHelper();
            String[] urls = {
                    "http://apache.org",
                    "www.apache.org",
                    "/temp",
                    "c:/temp",
                    "http://apache.org/default.php?s=isTramsformed&submit=Search&la=*&li=*"
            };

            for (int i = 0; i < urls.length; i++) {
                XSSFHyperlink link = helper.createHyperlink(HyperlinkType.URL);
                link.setAddress(urls[i]);
                Cell cell = row.createCell(i);
                cell.setCellValue(urls[i]);
                cell.setHyperlink(link);
            }

            workbook.write(outputStream);
        }

        extractXmlEntries(outputWorkbook, caseName, fixturesRoot);
    }

    @Test
    public void GenerateSheetLayoutFixture() throws Exception {
        Path fixturesRoot = getFixturesRoot();
        Path workbooksRoot = fixturesRoot.resolve("_workbooks");
        Files.createDirectories(fixturesRoot);
        Files.createDirectories(workbooksRoot);

        String caseName = "poi-integration-sheet-layout";
        cleanupCaseFiles(fixturesRoot, caseName);

        Path outputWorkbook = workbooksRoot.resolve(caseName + ".xlsx");
        try (XSSFWorkbook workbook = new XSSFWorkbook();
             OutputStream outputStream = Files.newOutputStream(outputWorkbook)) {
            setStableCoreProperties(workbook);

            Sheet sheet = workbook.createSheet("layout");
            for (int rowIndex = 0; rowIndex < 8; rowIndex++) {
                Row row = sheet.createRow(rowIndex);
                row.createCell(0).setCellValue("R" + (rowIndex + 1));
                row.createCell(1).setCellValue(rowIndex);
            }

            sheet.getRow(1).setHeightInPoints(24);
            sheet.setColumnWidth(0, 20 * 256);
            sheet.setColumnWidth(1, 14 * 256);
            sheet.setColumnHidden(2, true);
            sheet.addMergedRegion(CellRangeAddress.valueOf("A4:B5"));
            sheet.addMergedRegion(CellRangeAddress.valueOf("D1:E1"));
            sheet.createFreezePane(2, 4);
            sheet.groupRow(5, 7);

            workbook.write(outputStream);
        }

        extractXmlEntries(outputWorkbook, caseName, fixturesRoot);
    }

    @Test
    public void GenerateFormulaRecalculationFixture() throws Exception {
        Path fixturesRoot = getFixturesRoot();
        Path workbooksRoot = fixturesRoot.resolve("_workbooks");
        Files.createDirectories(fixturesRoot);
        Files.createDirectories(workbooksRoot);

        String caseName = "poi-integration-formula-recalculation";
        cleanupCaseFiles(fixturesRoot, caseName);

        Path outputWorkbook = workbooksRoot.resolve(caseName + ".xlsx");
        try (XSSFWorkbook workbook = new XSSFWorkbook();
             OutputStream outputStream = Files.newOutputStream(outputWorkbook)) {
            setStableCoreProperties(workbook);

            Sheet sheet = workbook.createSheet("formula");
            Row row = sheet.createRow(0);
            row.createCell(1).setCellValue(2);
            row.createCell(2).setCellValue(3);
            row.createCell(0).setCellFormula("B1+C1");
            workbook.setForceFormulaRecalculation(true);

            workbook.write(outputStream);
        }

        extractXmlEntries(outputWorkbook, caseName, fixturesRoot);
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

    private static Path getFixturesRoot() {
        return getRepoRoot()
                .resolve("tests")
                .resolve("DotnetPoi.Interop.Tests")
                .resolve("fixtures")
                .resolve("poi-integration");
    }

    private static void setStableCoreProperties(XSSFWorkbook workbook) throws Exception {
        workbook.getProperties().getCoreProperties().setCreated("2007-01-02T03:04:05Z");
        workbook.getProperties().getCoreProperties().setModified("2007-01-02T03:04:05Z");
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

    private static void extractXmlEntries(Path workbookPath, String caseName, Path fixturesRoot) throws IOException {
        try (ZipFile zipFile = new ZipFile(workbookPath.toFile())) {
            Enumeration<? extends ZipEntry> entries = zipFile.entries();
            while (entries.hasMoreElements()) {
                ZipEntry entry = entries.nextElement();
                String name = entry.getName();
                if (!isInterestingPackageEntry(name)) {
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

    private static boolean isInterestingPackageEntry(String name) {
        return name.endsWith(".xml")
                || name.endsWith(".rels")
                || name.endsWith(".vml")
                || name.endsWith(".bin");
    }
}
