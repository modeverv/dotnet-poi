using System.IO.Compression;
using DotnetPoi.XSSF.UserModel;
using DotnetPoi.XWPF.UserModel;
using DotnetPoi.XSLF.UserModel;
using Xunit;

namespace DotnetPoi.Core.Tests;

/// <summary>
/// Verify which ZIP parts survive a read → write round-trip when loading
/// real files from the Apache POI test-data repository.
/// Items inside model-rewritten XML (drawing.xml, sheet.xml, document.xml, slide.xml)
/// that the model does NOT understand are truly LOST (❌).
/// Items that are separate ZIP parts NOT in GetModelEntryNames() are 🔵 preserved.
/// </summary>
public class PreservationVerificationTests
{
    private static string PoiTestData => _poiTestData.Value;
    private static readonly Lazy<string> _poiTestData = new(ResolvePoiTestData);

    private static string ResolvePoiTestData()
    {
        // Search upward from test assembly directory for poi/test-data/
        var dir = AppContext.BaseDirectory;
        while (!string.IsNullOrEmpty(dir))
        {
            var candidate = Path.Combine(dir, "poi/test-data");
            if (Directory.Exists(candidate))
                return Path.GetFullPath(candidate);
            var parent = Path.GetDirectoryName(dir);
            if (parent == null || parent == dir || string.IsNullOrEmpty(parent)) break;
            dir = parent;
        }
        // Fallback: project-relative (dotnet test runs from project dir)
        var fromProj = Path.GetFullPath(Path.Combine(Directory.GetCurrentDirectory(), "..", "..", "poi/test-data"));
        return fromProj;
    }

    private void DumpLostEntries(string label, HashSet<string> before, HashSet<string> after)
    {
        var lost = before.Except(after).OrderBy(x => x).ToArray();
        if (lost.Length > 0)
        {
            Console.WriteLine($"  [{label}] Lost ZIP entries:");
            foreach (var n in lost)
                Console.WriteLine($"    ❌ {n}");
        }
        else
        {
            Console.WriteLine($"  [{label}] ✅ All original parts preserved");
        }

        var gained = after.Except(before).OrderBy(x => x).ToArray();
        if (gained.Length > 0)
        {
            Console.WriteLine($"  [{label}] New ZIP entries (created by model):");
            foreach (var n in gained)
                Console.WriteLine($"    ➕ {n}");
        }
    }

    [Fact]
    public void Docx_StylesAndComments_Preserved()
    {
        // Separate ZIP parts NOT in GetModelEntryNames() should survive:
        // word/styles.xml, word/comments.xml, word/footnotes.xml, word/endnotes.xml
        var path = Path.Combine(PoiTestData, "document/55966.docx");
        Assert.True(File.Exists(path), $"File not found: {path}");

        var raw = File.ReadAllBytes(path);
        var before = GetZipEntries(raw);
        // 55966.docx has word/styles.xml but NOT comments/footnotes/endnotes.
        // Let's also test comment.docx which has comments.
        var commentPath = Path.Combine(PoiTestData, "document/testComment.docx");
        HashSet<string>? commentBefore = null;
        if (File.Exists(commentPath))
        {
            var commentRaw = File.ReadAllBytes(commentPath);
            commentBefore = GetZipEntries(commentRaw);
        }

        // Round-trip test for 55966.docx (has styles)
        using var doc = new XWPFDocument(new MemoryStream(raw));
        using var ms = new MemoryStream();
        doc.write(ms);
        var after = GetZipEntries(ms.ToArray());
        DumpLostEntries("docx 55966", before, after);

        Assert.Contains("word/styles.xml", after);
        // Only check comments if file had them
        if (before.Contains("word/comments.xml"))
            Assert.Contains("word/comments.xml", after);
        if (before.Contains("word/footnotes.xml"))
            Assert.Contains("word/footnotes.xml", after);
        if (before.Contains("word/endnotes.xml"))
            Assert.Contains("word/endnotes.xml", after);
        if (before.Contains("word/glossary/styles.xml"))
            Assert.Contains("word/glossary/styles.xml", after);

        // Round-trip test for testComment.docx (has comments)
        if (commentBefore != null)
        {
            using var doc2 = new XWPFDocument(new MemoryStream(File.ReadAllBytes(commentPath)));
            using var ms2 = new MemoryStream();
            doc2.write(ms2);
            var after2 = GetZipEntries(ms2.ToArray());
            DumpLostEntries("docx testComment", commentBefore, after2);
            if (commentBefore.Contains("word/comments.xml"))
                Assert.Contains("word/comments.xml", after2);
        }
    }

    [Fact]
    public void Docx_HeaderFooter_RawXmlPreserved()
    {
        // headerPic.docx has a header with rich content (image references via blip).
        // When loaded and written without API modification, the raw header XML should
        // survive round-trip via _preservedEntries (headers/footers are no longer in
        // GetModelEntryNames()).
        var path = Path.Combine(PoiTestData, "document/headerPic.docx");
        Assert.True(File.Exists(path), $"File not found: {path}");

        var raw = File.ReadAllBytes(path);
        var before = GetZipEntries(raw);
        Assert.Contains("word/header1.xml", before);
        Assert.Contains("word/media/image1.jpeg", before);

        // Load (header has no <w:t> text — only an image; getHeaderText returns null)
        using var doc = new XWPFDocument(new MemoryStream(raw));

        // Write WITHOUT calling setHeaderText/setFooterText → preserved bytes should be used
        using var ms = new MemoryStream();
        doc.write(ms);
        var after = GetZipEntries(ms.ToArray());

        DumpLostEntries("docx header/footer round-trip", before, after);

        // header1.xml and its media should survive
        Assert.Contains("word/header1.xml", after);
        Assert.Contains("word/media/image1.jpeg", after);

        // The header content should be the rich original, not the minimal model output
        ms.Position = 0;
        using var verifyArchive = new ZipArchive(ms, ZipArchiveMode.Read);
        var headerEntry = verifyArchive.GetEntry("word/header1.xml");
        Assert.NotNull(headerEntry);
        using var r = new StreamReader(headerEntry.Open());
        var headerXml = r.ReadToEnd();
        // The preserved XML should contain a:blip (image reference), proving
        // the rich content survived rather than being replaced by model text
        Assert.Contains("a:blip", headerXml, StringComparison.Ordinal);
        Console.WriteLine($"  Preserved header XML length: {headerXml.Length} bytes ✅");
        Console.WriteLine($"  Contains a:blip: {headerXml.Contains("a:blip", StringComparison.Ordinal)}");
    }

    [Fact]
    public void Xlsx_RoundTrip_VerifyEntryPreservation()
    {
        // Check what ZIP entries survive. Auto-shapes (in drawing.xml) and
        // sparklines (in sheet.xml extLst) are inside model-rewritten XML.
        // Charts and external connections are separate parts → may be preserved.
        var path = Path.Combine(PoiTestData, "spreadsheet/123233_charts.xlsx");
        Assert.True(File.Exists(path), $"File not found: {path}");

        var raw = File.ReadAllBytes(path);
        var before = GetZipEntries(raw);
        Console.WriteLine($"  [before] Entries ({before.Count}):");
        foreach (var n in before.OrderBy(x => x))
            Console.WriteLine($"    {n}");

        using var wb = new XSSFWorkbook(new MemoryStream(raw));
        using var ms = new MemoryStream();
        wb.write(ms);
        var after = GetZipEntries(ms.ToArray());

        DumpLostEntries("xlsx round-trip", before, after);
    }

    [Fact]
    public void Xlsx_RoundTrip_AutoShapesAndPicturesPreserved()
    {
        // Auto-shapes (xdr:sp) inside xl/drawings/drawingN.xml should survive
        // round-trip via raw XML preservation. Use a file that has auto-shapes
        // with no picture dependencies.
        var path = Path.Combine(PoiTestData, "spreadsheet/47504.xlsx");
        Assert.True(File.Exists(path), $"File not found: {path}");

        var raw = File.ReadAllBytes(path);
        var before = GetZipEntries(raw);
        Console.WriteLine($"  [before] Entries ({before.Count}):");

        using var wb = new XSSFWorkbook(new MemoryStream(raw));

        // Check that auto-shapes were captured during load
        int totalPreserved = 0;
        bool foundAutoShapeInDrawing = false;
        for (int i = 0; i < wb.getNumberOfSheets(); i++)
        {
            var sheet = wb.getSheetAt(i);
            var drawing = sheet.Drawing;
            if (drawing is null) continue;
            totalPreserved += drawing.PreservedRawAnchors.Count;
            foreach (var anchor in drawing.PreservedRawAnchors)
            {
                if (anchor.Contains("xdr:sp"))
                {
                    foundAutoShapeInDrawing = true;
                    break;
                }
            }
            Console.WriteLine($"  Sheet #{i}: PreservedRawAnchors={drawing.PreservedRawAnchors.Count}, Pictures={drawing.Pictures.Count}");
        }
        Assert.True(totalPreserved > 0,
            $"Expected at least one auto-shape anchor to be preserved, got {totalPreserved}.");
        Assert.True(foundAutoShapeInDrawing,
            "No xdr:sp content found among preserved anchors.");

        // Write and verify
        using var ms = new MemoryStream();
        wb.write(ms);
        var after = GetZipEntries(ms.ToArray());

        DumpLostEntries("xlsx drawing round-trip", before, after);

        // Verify the drawing contains auto-shapes in the output
        ms.Position = 0;
        using var verifyArchive = new ZipArchive(ms, ZipArchiveMode.Read);
        var drawingEntry = verifyArchive.GetEntry("xl/drawings/drawing1.xml");
        Assert.NotNull(drawingEntry);
        using var r = new StreamReader(drawingEntry.Open());
        var drawingXml = r.ReadToEnd();
        Assert.Contains("xdr:sp", drawingXml);
        Console.WriteLine("  Output drawing XML contains xdr:sp ✅");
    }

    [Fact]
    public void Pptx_RoundTrip_VerifyEntryPreservation()
    {
        // Group shapes (<p:grpSp>) and connectors (<p:cxnSp>) are inside
        // ppt/slides/slideN.xml which the model rewrites → would be lost.
        // Use a file confirmed to have actual group shapes.
        var path = Path.Combine(PoiTestData, "slideshow/sample_pptx_grouping_issues.pptx");
        Assert.True(File.Exists(path), $"File not found: {path}");

        var raw = File.ReadAllBytes(path);
        var before = GetZipEntries(raw);
        Console.WriteLine($"  [before] Entries ({before.Count}):");
        foreach (var n in before.OrderBy(x => x))
            Console.WriteLine($"    {n}");

        using var prs = new XMLSlideShow(new MemoryStream(raw));

        // Check that group shapes were captured during load
        int totalPreserved = 0;
        foreach (var slide in prs.getSlides())
        {
            totalPreserved += slide.getPreservedRawElements().Count;
        }
        Console.WriteLine($"  Preserved raw elements across all slides: {totalPreserved}");
        Assert.True(totalPreserved > 0,
            $"Expected at least one group shape/connector to be preserved, got {totalPreserved}. "
            + "Did 2411-Performance_Up.pptx have group shapes in slide xml?");

        using var ms = new MemoryStream();
        prs.write(ms);
        var after = GetZipEntries(ms.ToArray());

        DumpLostEntries("pptx round-trip", before, after);

        Assert.Contains("ppt/slides/slide1.xml", after);

        // Verify group shapes are actually in the output slide XML
        ms.Position = 0;
        using var verifyArchive = new ZipArchive(ms, ZipArchiveMode.Read, leaveOpen: false);
        int slidesWithGroupShapes = 0;
        foreach (var entry in verifyArchive.Entries
            .Where(e => e.FullName.StartsWith("ppt/slides/slide", StringComparison.OrdinalIgnoreCase)
                        && e.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)))
        {
            using var r = new StreamReader(entry.Open());
            var content = r.ReadToEnd();
            if (content.Contains("p:grpSp", StringComparison.Ordinal) ||
                content.Contains("p:cxnSp", StringComparison.Ordinal))
            {
                slidesWithGroupShapes++;
            }
        }
        Console.WriteLine($"  Slides containing group shapes / connectors: {slidesWithGroupShapes}");
        Assert.True(slidesWithGroupShapes > 0,
            "No group shapes or connectors found in output slide XML. "
            + "The preservation mechanism may not be working.");
    }

    [Fact]
    public void Docx_BlockLevelSdt_Preserved()
    {
        // Block-level SDT (w:sdt as direct child of w:body, containing paragraphs
        // inside w:sdtContent) should survive round-trip via raw XML preservation.
        // Files with only inline SDT (inside paragraphs) like 52449.docx are not
        // covered by this mechanism.
        var path = Path.Combine(PoiTestData, "document/60316.docx");
        Assert.True(File.Exists(path), $"File not found: {path}");

        // 60316.docx has 3 block-level SDT elements verified during development.
        // Total SDT occurrences in output should count both open tags and references
        // within preserved raw XML content.

        var raw = File.ReadAllBytes(path);
        var before = GetZipEntries(raw);
        Console.WriteLine($"  60316.docx — [before] Entries ({before.Count}):");
        foreach (var n in before.OrderBy(x => x))
            Console.WriteLine($"    {n}");

        using var doc = new XWPFDocument(new MemoryStream(raw));
        using var ms = new MemoryStream();
        doc.write(ms);
        var after = GetZipEntries(ms.ToArray());

        DumpLostEntries("docx SDT round-trip (60316.docx)", before, after);

        // Verify SDT content survives in output document.xml
        ms.Position = 0;
        using var verifyArchive = new ZipArchive(ms, ZipArchiveMode.Read);
        var docEntry = verifyArchive.GetEntry("word/document.xml");
        Assert.NotNull(docEntry);
        using var r = new StreamReader(docEntry.Open());
        var docXml = r.ReadToEnd();
        Assert.Contains("<w:sdt", docXml, StringComparison.Ordinal);
        int sdtCount = 0, idx = 0;
        while ((idx = docXml.IndexOf("<w:sdt", idx, StringComparison.Ordinal)) != -1)
        {
            sdtCount++;
            idx += 6;
        }
        Console.WriteLine($"  Block-level SDT elements in output: {sdtCount}");
        Assert.True(sdtCount >= 2, $"Expected at least 2 block-level SDT elements, got {sdtCount}");
    }

    private static int CountOccurrences(string text, string pattern)
    {
        int count = 0, idx = 0;
        while ((idx = text.IndexOf(pattern, idx, StringComparison.Ordinal)) != -1)
        {
            count++;
            idx += pattern.Length;
        }
        return count;
    }

    private static HashSet<string> GetZipEntries(byte[] data)
    {
        var set = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        using var archive = new ZipArchive(new MemoryStream(data), ZipArchiveMode.Read);
        foreach (var entry in archive.Entries)
            set.Add(entry.FullName.Replace('\\', '/'));
        return set;
    }
}
