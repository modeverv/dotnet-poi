package org.dotnetpoi.interop;

import org.apache.poi.ooxml.POIXMLTypeLoader;
import org.apache.xmlbeans.XmlObject;
import org.apache.xmlbeans.XmlOptions;
import org.junit.jupiter.api.Test;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCell;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCellFormula;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTRow;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTSst;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTWorkbook;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTWorksheet;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.STCellType;

import javax.xml.namespace.QName;
import java.io.IOException;
import java.io.OutputStream;
import java.nio.file.DirectoryStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.LinkedHashMap;
import java.util.Map;

public class XmlBeansOutputProbeTest {
    private static final String MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private static final String REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private static final String CASE_PREFIX = "xmlbeans-";

    @Test
    public void GenerateXmlBeansOutputFixtures() throws Exception {
        Path outputRoot = getOutputRoot();
        Files.createDirectories(outputRoot);
        cleanup(outputRoot);

        for (Map.Entry<String, XmlObject> sample : createSamples().entrySet()) {
            writeSample(outputRoot, sample.getKey(), sample.getValue(), null, "default-save");
            writeSample(outputRoot, sample.getKey(), sample.getValue(), POIXMLTypeLoader.DEFAULT_XML_OPTIONS, "poi-options");
        }
    }

    private static Map<String, XmlObject> createSamples() {
        Map<String, XmlObject> samples = new LinkedHashMap<>();
        samples.put("worksheet-empty-sheet-data", createEmptySheetDataWorksheet());
        samples.put("worksheet-zero-cell", createZeroCellWorksheet());
        samples.put("worksheet-formula-empty-result", createFormulaWorksheet());
        samples.put("workbook-two-sheets", createWorkbookWithTwoSheets());
        samples.put("shared-strings-escaping", createSharedStrings());
        return samples;
    }

    private static CTWorksheet createEmptySheetDataWorksheet() {
        CTWorksheet worksheet = CTWorksheet.Factory.newInstance();
        worksheet.addNewSheetData();
        worksheet.addNewPageMargins()
                .setBottom(0.75);
        worksheet.getPageMargins().setFooter(0.3);
        worksheet.getPageMargins().setHeader(0.3);
        worksheet.getPageMargins().setLeft(0.7);
        worksheet.getPageMargins().setRight(0.7);
        worksheet.getPageMargins().setTop(0.75);
        return worksheet;
    }

    private static CTWorksheet createZeroCellWorksheet() {
        CTWorksheet worksheet = CTWorksheet.Factory.newInstance();
        CTRow row = worksheet.addNewSheetData().addNewRow();
        row.setR(1);

        CTCell cell = row.addNewC();
        cell.setR("A1");
        cell.setV("0");
        return worksheet;
    }

    private static CTWorksheet createFormulaWorksheet() {
        CTWorksheet worksheet = CTWorksheet.Factory.newInstance();
        CTRow row = worksheet.addNewSheetData().addNewRow();
        row.setR(1);

        CTCell formulaCell = row.addNewC();
        formulaCell.setR("A1");
        formulaCell.setT(STCellType.STR);
        CTCellFormula formula = formulaCell.addNewF();
        formula.setStringValue("SUM(1,2,3)");
        formulaCell.setV("");
        return worksheet;
    }

    private static CTWorkbook createWorkbookWithTwoSheets() {
        CTWorkbook workbook = CTWorkbook.Factory.newInstance();
        workbook.addNewWorkbookPr().setDate1904(false);

        var sheets = workbook.addNewSheets();
        var sheet1 = sheets.addNewSheet();
        sheet1.setName("Alpha");
        sheet1.setSheetId(1);
        sheet1.setId("rId1");

        var sheet2 = sheets.addNewSheet();
        sheet2.setName("Beta & Co");
        sheet2.setSheetId(2);
        sheet2.setId("rId2");

        return workbook;
    }

    private static CTSst createSharedStrings() {
        CTSst sst = CTSst.Factory.newInstance();
        sst.setCount(2);
        sst.setUniqueCount(2);

        sst.addNewSi().setT("A&B <C> \"quoted\" 'single'");

        sst.addNewSi().setT("");

        return sst;
    }

    private static void writeSample(Path outputRoot,
                                    String caseName,
                                    XmlObject xmlObject,
                                    XmlOptions options,
                                    String mode) throws IOException {
        XmlOptions effectiveOptions = options == null ? new XmlOptions() : new XmlOptions(options);
        effectiveOptions.setSaveSyntheticDocumentElement(syntheticDocumentElement(xmlObject));

        Path outputPath = outputRoot.resolve(CASE_PREFIX + caseName + "__" + mode + ".xml");
        try (OutputStream outputStream = Files.newOutputStream(outputPath)) {
            xmlObject.save(outputStream, effectiveOptions);
        }
    }

    private static QName syntheticDocumentElement(XmlObject xmlObject) {
        if (xmlObject instanceof CTWorksheet) {
            return new QName(MAIN_NS, "worksheet");
        }
        if (xmlObject instanceof CTWorkbook) {
            return new QName(MAIN_NS, "workbook");
        }
        if (xmlObject instanceof CTSst) {
            return new QName(MAIN_NS, "sst");
        }
        throw new IllegalArgumentException("Unsupported XMLBeans probe type: " + xmlObject.getClass().getName());
    }

    private static Path getOutputRoot() {
        Path directory = Paths.get("").toAbsolutePath();
        while (directory != null) {
            if (Files.exists(directory.resolve("DotnetPOI.sln"))) {
                return directory.resolve("tests")
                        .resolve("DotnetPoi.Interop.Tests")
                        .resolve("fixtures")
                        .resolve("xmlbeans-output");
            }
            directory = directory.getParent();
        }
        throw new IllegalStateException("Could not locate the dotnet-poi repository root.");
    }

    private static void cleanup(Path outputRoot) throws IOException {
        try (DirectoryStream<Path> stream = Files.newDirectoryStream(outputRoot, CASE_PREFIX + "*")) {
            for (Path path : stream) {
                Files.deleteIfExists(path);
            }
        }
    }
}
