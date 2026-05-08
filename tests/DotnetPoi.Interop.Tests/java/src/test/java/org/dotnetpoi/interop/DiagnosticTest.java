package org.dotnetpoi.interop;

import org.junit.jupiter.api.Test;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.extractor.WordExtractor;
import org.apache.poi.hwpf.model.FileInformationBlock;

import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import static org.junit.jupiter.api.Assertions.*;

public class DiagnosticTest {

    @Test
    void diagPhaseDoc() throws Exception {
        // Phase 13 no-op
        Path noop = findRepoRoot().resolve("tests/DotnetPoi.Interop.Tests/fixtures/from-dotnet-poi/phase13-noop-sample.doc");
        try (InputStream in = Files.newInputStream(noop);
             HWPFDocument doc = new HWPFDocument(in)) {
            System.out.println("=== Phase13 NoOp ===");
            String txt = doc.getDocumentText();
            System.out.println("  getDocumentText(): '" + escape(txt) + "'");
            System.out.println("  getDocumentText().length: " + txt.length());
            System.out.println("  fcClx: " + doc.getFileInformationBlock().getFcClx());
            System.out.println("  lcbClx: " + doc.getFileInformationBlock().getLcbClx());
            System.out.println("  fcPlcfbteChpx: " + doc.getFileInformationBlock().getFcPlcfbteChpx());
            System.out.println("  lcbPlcfbteChpx: " + doc.getFileInformationBlock().getLcbPlcfbteChpx());
            System.out.println("  fcPlcfbtePapx: " + doc.getFileInformationBlock().getFcPlcfbtePapx());
            System.out.println("  lcbPlcfbtePapx: " + doc.getFileInformationBlock().getLcbPlcfbtePapx());
            System.out.println("  fcStshf: " + doc.getFileInformationBlock().getFcStshf());
            System.out.println("  lcbStshf: " + doc.getFileInformationBlock().getLcbStshf());

            WordExtractor ex = new WordExtractor(doc);
            String et = ex.getText();
            System.out.println("  extractor text: '" + escape(et) + "'");
        }

        // Phase 14 edited
        Path edited = findRepoRoot().resolve("tests/DotnetPoi.Interop.Tests/fixtures/from-dotnet-poi/phase14-edited-sample.doc");
        try (InputStream in = Files.newInputStream(edited);
             HWPFDocument doc = new HWPFDocument(in)) {
            System.out.println("=== Phase14 Edited ===");
            String txt = doc.getDocumentText();
            System.out.println("  getDocumentText(): '" + escape(txt) + "'");
            System.out.println("  getDocumentText().length: " + txt.length());
            System.out.println("  fcClx: " + doc.getFileInformationBlock().getFcClx());
            System.out.println("  lcbClx: " + doc.getFileInformationBlock().getLcbClx());
            System.out.println("  fcPlcfbteChpx: " + doc.getFileInformationBlock().getFcPlcfbteChpx());
            System.out.println("  lcbPlcfbteChpx: " + doc.getFileInformationBlock().getLcbPlcfbteChpx());
            System.out.println("  fcPlcfbtePapx: " + doc.getFileInformationBlock().getFcPlcfbtePapx());
            System.out.println("  lcbPlcfbtePapx: " + doc.getFileInformationBlock().getLcbPlcfbtePapx());
            System.out.println("  fcStshf: " + doc.getFileInformationBlock().getFcStshf());
            System.out.println("  lcbStshf: " + doc.getFileInformationBlock().getLcbStshf());

            // Try range.text()
            try {
                org.apache.poi.hwpf.usermodel.Range range = doc.getRange();
                System.out.println("  getRange(): start=" + range.getStartOffset() + " end=" + range.getEndOffset());
                System.out.println("  numParagraphs: " + range.numParagraphs());
                if (range.numParagraphs() > 0) {
                    org.apache.poi.hwpf.usermodel.Paragraph p0 = range.getParagraph(0);
                    System.out.println("  P0: start=" + p0.getStartOffset() + " end=" + p0.getEndOffset() + " text='" + escape(p0.text()) + "'");
                }
                String rt = range.text();
                System.out.println("  range.text(): '" + escape(rt) + "'");
            } catch (Exception e) {
                System.out.println("  range.text() FAILED: " + e.getClass().getName() + ": " + e.getMessage());
                e.printStackTrace(System.out);
            }

            WordExtractor ex = new WordExtractor(doc);
            String et = ex.getText();
            System.out.println("  extractor text: '" + escape(et) + "'");
        }
    }

    private static String escape(String s) {
        if (s == null) return "(null)";
        if (s.isEmpty()) return "(empty)";
        return s.replace("\r", "\\r").replace("\n", "\\n").replace("\f", "\\f").substring(0, Math.min(s.length(), 200));
    }

    private static Path findRepoRoot() {
        Path p = Path.of(".").toAbsolutePath().normalize();
        while (p != null && !Files.exists(p.resolve("DotnetPOI.sln"))) {
            p = p.getParent();
        }
        return p;
    }
}
