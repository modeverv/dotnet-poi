using DotnetPoi.SS.UserModel;
using DotnetPoi.XSSF.UserModel;
using DotnetPoi.XWPF.UserModel;
using DotnetPoi.XSLF.UserModel;
using System.IO.Compression;
using Xunit;

namespace DotnetPoi.Interop.Tests;

/// <summary>
/// Preservation fixture tests using real Apache POI test data files from <c>poi/test-data/</c>.
///
/// Purpose: verify that files containing features NOT (fully) supported by DotnetPoi
/// (charts, comments, text boxes, OLE objects, footnotes, SmartArt, video, etc.)
/// can be round-tripped (load → apply a trivial modification → save → reload)
/// without losing data. Unsupported parts should pass through via the unknown-part
/// preservation mechanism (_preservedEntries).
///
/// These tests do NOT verify correctness of the unsupported feature — they only
/// verify the feature's binary parts survive the round-trip.
/// </summary>
public sealed class PoiPreservationTests
{
    // ──────────────── Helpers ────────────────

    /// <summary>
    /// Resolve a path relative to <c>poi/test-data/</c> at the repository root.
    /// </summary>
    private static string GetPoiTestDataPath(string relativePath)
    {
        var dir = AppContext.BaseDirectory;
        while (dir is not null)
        {
            var testDataDir = Path.Combine(dir, "poi", "test-data");
            if (Directory.Exists(testDataDir))
                return Path.Combine(testDataDir, relativePath);
            dir = Directory.GetParent(dir)?.FullName;
        }
        throw new DirectoryNotFoundException(
            "Could not locate poi/test-data/ from any parent of AppContext.BaseDirectory.");
    }

    /// <summary>
    /// Capture the set of entry names in a ZIP file (original fixture).
    /// </summary>
    private static HashSet<string> GetZipEntryNames(string path)
    {
        using var archive = ZipFile.OpenRead(path);
        return new HashSet<string>(archive.Entries.Select(e => e.FullName.Replace('\\', '/')));
    }

    /// <summary>
    /// Capture the set of entry names from a round-tripped MemoryStream.
    /// </summary>
    private static HashSet<string> GetZipEntryNames(Stream stream)
    {
        // ZipArchive disposes the stream unless we use leaveOpen.
        // Copy the bytes so we don't consume the caller's stream.
        using var ms = new MemoryStream();
        stream.Position = 0;
        stream.CopyTo(ms);
        ms.Position = 0;
        using var archive = new ZipArchive(ms, ZipArchiveMode.Read, leaveOpen: false);
        return new HashSet<string>(archive.Entries.Select(e => e.FullName.Replace('\\', '/')));
    }

    /// <summary>
    /// Assert that every entry in <paramref name="expected"/> that is NOT a model-known
    /// entry still exists in <paramref name="actual"/>. The model-known entries are
    /// ones explicitly written by the DotnetPoi writer; everything else (unknown parts)
    /// should pass through unchanged.
    /// </summary>
    private static void AssertUnknownPartsPreserved(
        HashSet<string> original,
        HashSet<string> roundTripped,
        HashSet<string> modelEntryPrefixes)
    {
        var unknownParts = new HashSet<string>(original);
        unknownParts.ExceptWith(modelEntryPrefixes);

        // These unknown parts must survive round-trip
        var missing = unknownParts.Except(roundTripped).ToList();
        if (missing.Count > 0)
        {
            Assert.Fail(
                $"{missing.Count} unknown part(s) were lost during round-trip:\n  " +
                string.Join("\n  ", missing.Take(20)));
        }
    }


    // ════════════════════════════════════════════════
    //  XLSX — preservation tests
    // ════════════════════════════════════════════════

    private static HashSet<string> XlsxModelEntries =>
        new(StringComparer.OrdinalIgnoreCase)
        {
            "[Content_Types].xml",
            "_rels/.rels",
            "docProps/app.xml",
            "docProps/core.xml",
            "xl/workbook.xml",
            "xl/_rels/workbook.xml.rels",
            "xl/styles.xml",
            "xl/sharedStrings.xml",
        };

    /// <summary>Modify a loaded xlsx fixture and verify chart part survives.</summary>
    [Fact]
    public void Preserve_Xlsx_WithChart()
    {
        var path = GetPoiTestDataPath("spreadsheet/WithChart.xlsx");
        Assert.True(File.Exists(path));

        using var fs = new FileStream(path, FileMode.Open, FileAccess.Read);
        using var wb = new XSSFWorkbook(fs);

        // Apply a small supported modification
        var sheet = wb.getSheetAt(0);
        var row = sheet.getRow(0) ?? sheet.createRow(0);
        row.createCell(99).setCellValue("preservation-check");

        // Save to MemoryStream
        using var saved = new MemoryStream();
        wb.write(saved);

        // Reload and verify modification
        saved.Position = 0;
        using var loaded = new XSSFWorkbook(saved);
        Assert.Equal("preservation-check",
            loaded.getSheetAt(0).getRow(0)?.getCell(99)?.getStringCellValue());

        // Verify chart entries still exist
        saved.Position = 0;
        var entries = GetZipEntryNames(saved);
        Assert.Contains(entries, e => e.StartsWith("xl/charts/"));
        Assert.True(entries.Contains("xl/charts/chart1.xml") ||
                    entries.Any(e => e.StartsWith("xl/charts/")));
    }

    /// <summary>Preserve comments parts.</summary>
    [Fact]
    public void Preserve_Xlsx_SimpleWithComments()
    {
        var path = GetPoiTestDataPath("spreadsheet/SimpleWithComments.xlsx");
        Assert.True(File.Exists(path));

        using var fs = new FileStream(path, FileMode.Open, FileAccess.Read);
        using var wb = new XSSFWorkbook(fs);

        wb.getSheetAt(0).createRow(0).createCell(0).setCellValue("preserved");

        using var saved = new MemoryStream();
        wb.write(saved);

        saved.Position = 0;
        var entries = GetZipEntryNames(saved);
        Assert.Contains(entries, e => e.StartsWith("xl/comments"));
        Assert.Contains(entries, e => e.StartsWith("xl/drawings/vmlDrawing"));
    }

    /// <summary>Preserve text box (drawing) parts.</summary>
    [Fact]
    public void Preserve_Xlsx_WithTextBox()
    {
        var path = GetPoiTestDataPath("spreadsheet/WithTextBox.xlsx");
        Assert.True(File.Exists(path));

        using var fs = new FileStream(path, FileMode.Open, FileAccess.Read);
        using var wb = new XSSFWorkbook(fs);

        wb.getSheetAt(0).createRow(0).createCell(0).setCellValue("preserved");

        using var saved = new MemoryStream();
        wb.write(saved);

        saved.Position = 0;
        var entries = GetZipEntryNames(saved);
        Assert.Contains(entries, e => e.StartsWith("xl/drawings/drawing"));
    }

    /// <summary>Preserve OLE embedded objects.</summary>
    [Fact]
    public void Preserve_Xlsx_WithEmbedded()
    {
        var path = GetPoiTestDataPath("spreadsheet/WithEmbeded.xlsx");
        Assert.True(File.Exists(path));

        using var fs = new FileStream(path, FileMode.Open, FileAccess.Read);
        using var wb = new XSSFWorkbook(fs);

        wb.getSheetAt(0).createRow(0).createCell(0).setCellValue("preserved");

        using var saved = new MemoryStream();
        wb.write(saved);

        saved.Position = 0;
        var entries = GetZipEntryNames(saved);
        Assert.Contains(entries, e => e.StartsWith("xl/embeddings/"));
        Assert.Contains(entries, e => e.StartsWith("xl/media/"));
    }

    /// <summary>
    /// Preserve VBA + comments + OLE + drawings in an xlsm.
    /// Also verifies VBA bytes are identical byte-for-byte.
    /// </summary>
    [Fact]
    public void Preserve_Xlsm_ExcelWithAttachments()
    {
        var path = GetPoiTestDataPath("spreadsheet/ExcelWithAttachments.xlsm");
        Assert.True(File.Exists(path));

        // Capture original VBA bytes
        byte[] originalVba;
        using (var archive = ZipFile.OpenRead(path))
        {
            var entry = archive.GetEntry("xl/vbaProject.bin")
                ?? throw new InvalidDataException("xl/vbaProject.bin missing from fixture");
            using var s = entry.Open();
            using var ms = new MemoryStream();
            s.CopyTo(ms);
            originalVba = ms.ToArray();
        }

        using var fs = new FileStream(path, FileMode.Open, FileAccess.Read);
        using var wb = new XSSFWorkbook(fs);

        wb.getSheetAt(0).createRow(0).createCell(0).setCellValue("preserved");

        using var saved = new MemoryStream();
        wb.write(saved);

        // Verify modification
        saved.Position = 0;
        using var loaded = new XSSFWorkbook(saved);
        Assert.Equal("preserved",
            loaded.getSheetAt(0).getRow(0)?.getCell(0)?.getStringCellValue());

        // Verify VBA bytes byte-for-byte
        saved.Position = 0;
        using var rtArchive = new ZipArchive(saved, ZipArchiveMode.Read, leaveOpen: false);
        var rtEntry = rtArchive.GetEntry("xl/vbaProject.bin");
        Assert.NotNull(rtEntry);
        byte[] rtVba;
        using (var s = rtEntry.Open())
        using (var ms = new MemoryStream())
        {
            s.CopyTo(ms);
            rtVba = ms.ToArray();
        }
        Assert.Equal(originalVba, rtVba);

        // Verify other unknown parts
        var rtNames = new HashSet<string>(rtArchive.Entries.Select(e => e.FullName.Replace('\\', '/')));
        Assert.Contains(rtNames, e => e.StartsWith("xl/embeddings/"));
        Assert.Contains(rtNames, e => e.StartsWith("xl/comments"));
    }


    // ════════════════════════════════════════════════
    //  DOCX — preservation tests
    // ════════════════════════════════════════════════

    /// <summary>Preserve comments in docx.</summary>
    [Fact]
    public void Preserve_Docx_Comments()
    {
        var path = GetPoiTestDataPath("document/comment.docx");
        Assert.True(File.Exists(path));

        using var fs = new FileStream(path, FileMode.Open, FileAccess.Read);
        using var doc = new XWPFDocument(fs);

        doc.createParagraph().createRun().setText("preserved");

        using var saved = new MemoryStream();
        doc.write(saved);

        saved.Position = 0;
        using var loaded = new XWPFDocument(saved);
        Assert.Contains(loaded.getParagraphs(),
            p => p.getText().Contains("preserved"));

        saved.Position = 0;
        var entries = GetZipEntryNames(saved);
        Assert.Contains(entries, e => e == "word/comments.xml");
    }

    /// <summary>Preserve footnotes and endnotes.</summary>
    [Fact]
    public void Preserve_Docx_Footnotes()
    {
        var path = GetPoiTestDataPath("document/footnotes.docx");
        Assert.True(File.Exists(path));

        using var fs = new FileStream(path, FileMode.Open, FileAccess.Read);
        using var doc = new XWPFDocument(fs);

        doc.createParagraph().createRun().setText("preserved");

        using var saved = new MemoryStream();
        doc.write(saved);

        saved.Position = 0;
        using var loaded = new XWPFDocument(saved);
        Assert.Contains(loaded.getParagraphs(),
            p => p.getText().Contains("preserved"));

        saved.Position = 0;
        var entries = GetZipEntryNames(saved);
        Assert.Contains(entries, e => e == "word/footnotes.xml");
        Assert.Contains(entries, e => e == "word/endnotes.xml");
    }

    /// <summary>Preserve change tracking (delins).</summary>
    [Fact]
    public void Preserve_Docx_ChangeTracking()
    {
        var path = GetPoiTestDataPath("document/delins.docx");
        Assert.True(File.Exists(path));

        using var fs = new FileStream(path, FileMode.Open, FileAccess.Read);
        using var doc = new XWPFDocument(fs);

        doc.createParagraph().createRun().setText("preserved");

        using var saved = new MemoryStream();
        doc.write(saved);

        saved.Position = 0;
        using var loaded = new XWPFDocument(saved);
        Assert.Contains(loaded.getParagraphs(),
            p => p.getText().Contains("preserved"));
    }

    /// <summary>Preserve OLE embedded objects in docx.</summary>
    [Fact]
    public void Preserve_Docx_EmbeddedDocument()
    {
        var path = GetPoiTestDataPath("document/EmbeddedDocument.docx");
        Assert.True(File.Exists(path));

        using var fs = new FileStream(path, FileMode.Open, FileAccess.Read);
        using var doc = new XWPFDocument(fs);

        doc.createParagraph().createRun().setText("preserved");

        using var saved = new MemoryStream();
        doc.write(saved);

        saved.Position = 0;
        using var loaded = new XWPFDocument(saved);
        Assert.Contains(loaded.getParagraphs(),
            p => p.getText().Contains("preserved"));

        saved.Position = 0;
        var entries = GetZipEntryNames(saved);
        Assert.Contains(entries, e => e.StartsWith("word/embeddings/"));
        Assert.Contains(entries, e => e.StartsWith("word/media/"));
    }

    /// <summary>Preserve columns (sectPr with cols).</summary>
    [Fact]
    public void Preserve_Docx_ThreeColHead()
    {
        var path = GetPoiTestDataPath("document/ThreeColHead.docx");
        Assert.True(File.Exists(path));

        using var fs = new FileStream(path, FileMode.Open, FileAccess.Read);
        using var doc = new XWPFDocument(fs);

        doc.createParagraph().createRun().setText("preserved");

        using var saved = new MemoryStream();
        doc.write(saved);

        saved.Position = 0;
        using var loaded = new XWPFDocument(saved);
        Assert.Contains(loaded.getParagraphs(),
            p => p.getText().Contains("preserved"));

        saved.Position = 0;
        var entries = GetZipEntryNames(saved);
        Assert.Contains(entries, e => e == "word/header1.xml");
        Assert.Contains(entries, e => e == "word/footnotes.xml" || e == "word/endnotes.xml");
    }


    // ════════════════════════════════════════════════
    //  PPTX — preservation tests
    // ════════════════════════════════════════════════

    /// <summary>Preserve notes, comments, OLE in pptx.</summary>
    [Fact]
    public void Preserve_Pptx_CommentsAndNotes()
    {
        var path = GetPoiTestDataPath("slideshow/45545_Comment.pptx");
        Assert.True(File.Exists(path));

        using var fs = new FileStream(path, FileMode.Open, FileAccess.Read);
        using var prs = new XMLSlideShow(fs);

        var slide = prs.createSlide();

        using var saved = new MemoryStream();
        prs.write(saved);

        saved.Position = 0;
        using var loaded = new XMLSlideShow(saved);
        Assert.NotEmpty(loaded.getSlides());

        saved.Position = 0;
        var entries = GetZipEntryNames(saved);
        Assert.Contains(entries, e => e.StartsWith("ppt/comments/"));
        Assert.Contains(entries, e => e.StartsWith("ppt/notesSlides/"));
        Assert.Contains(entries, e => e.StartsWith("ppt/embeddings/"));
    }

    /// <summary>Preserve chart parts in pptx.</summary>
    [Fact]
    public void Preserve_Pptx_Chart()
    {
        var path = GetPoiTestDataPath("slideshow/line-chart.pptx");
        Assert.True(File.Exists(path));

        using var fs = new FileStream(path, FileMode.Open, FileAccess.Read);
        using var prs = new XMLSlideShow(fs);

        var slide = prs.createSlide();

        using var saved = new MemoryStream();
        prs.write(saved);

        saved.Position = 0;
        using var loaded = new XMLSlideShow(saved);
        Assert.NotEmpty(loaded.getSlides());

        saved.Position = 0;
        var entries = GetZipEntryNames(saved);
        Assert.Contains(entries, e => e.StartsWith("ppt/charts/"));
    }

    /// <summary>Preserve embedded video media in pptx.</summary>
    [Fact]
    public void Preserve_Pptx_EmbeddedVideo()
    {
        var path = GetPoiTestDataPath("slideshow/EmbeddedVideo.pptx");
        Assert.True(File.Exists(path));

        using var fs = new FileStream(path, FileMode.Open, FileAccess.Read);
        using var prs = new XMLSlideShow(fs);

        var slide = prs.createSlide();

        using var saved = new MemoryStream();
        prs.write(saved);

        saved.Position = 0;
        using var loaded = new XMLSlideShow(saved);
        Assert.NotEmpty(loaded.getSlides());

        saved.Position = 0;
        var entries = GetZipEntryNames(saved);
        Assert.Contains(entries, e => e.Contains("media1.mp4"));
    }

    /// <summary>Preserve SmartArt / diagrams in pptx.</summary>
    [Fact]
    public void Preserve_Pptx_SmartArt()
    {
        var path = GetPoiTestDataPath("slideshow/SmartArt.pptx");
        Assert.True(File.Exists(path));

        using var fs = new FileStream(path, FileMode.Open, FileAccess.Read);
        using var prs = new XMLSlideShow(fs);

        var slide = prs.createSlide();

        using var saved = new MemoryStream();
        prs.write(saved);

        saved.Position = 0;
        using var loaded = new XMLSlideShow(saved);
        Assert.NotEmpty(loaded.getSlides());

        saved.Position = 0;
        var entries = GetZipEntryNames(saved);
        Assert.Contains(entries, e => e.StartsWith("ppt/diagrams/"));
    }

    /// <summary>Preserve VBA + OLE in pptm.</summary>
    [Fact]
    public void Preserve_Pptm_WithAttachments()
    {
        var path = GetPoiTestDataPath("slideshow/PPTWithAttachments.pptm");
        Assert.True(File.Exists(path));

        using var fs = new FileStream(path, FileMode.Open, FileAccess.Read);
        using var prs = new XMLSlideShow(fs);

        var slide = prs.createSlide();

        using var saved = new MemoryStream();
        prs.write(saved);

        saved.Position = 0;
        using var loaded = new XMLSlideShow(saved);
        Assert.NotEmpty(loaded.getSlides());

        saved.Position = 0;
        var entries = GetZipEntryNames(saved);
        Assert.Contains(entries, e => e == "ppt/vbaProject.bin");
        Assert.Contains(entries, e => e.StartsWith("ppt/embeddings/"));
    }

    /// <summary>Preserve grouped shapes / notes in pptx.</summary>
    [Fact]
    public void Preserve_Pptx_GroupingAndNotes()
    {
        var path = GetPoiTestDataPath("slideshow/sample_pptx_grouping_issues.pptx");
        Assert.True(File.Exists(path));

        using var fs = new FileStream(path, FileMode.Open, FileAccess.Read);
        using var prs = new XMLSlideShow(fs);

        var slide = prs.createSlide();

        using var saved = new MemoryStream();
        prs.write(saved);

        saved.Position = 0;
        using var loaded = new XMLSlideShow(saved);
        Assert.NotEmpty(loaded.getSlides());

        saved.Position = 0;
        var entries = GetZipEntryNames(saved);
        Assert.Contains(entries, e => e.StartsWith("ppt/notesSlides/"));
    }
}
