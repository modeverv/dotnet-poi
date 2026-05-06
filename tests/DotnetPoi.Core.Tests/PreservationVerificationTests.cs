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

    private static HashSet<string> GetZipEntries(byte[] data)
    {
        var set = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        using var archive = new ZipArchive(new MemoryStream(data), ZipArchiveMode.Read);
        foreach (var entry in archive.Entries)
            set.Add(entry.FullName.Replace('\\', '/'));
        return set;
    }
}
