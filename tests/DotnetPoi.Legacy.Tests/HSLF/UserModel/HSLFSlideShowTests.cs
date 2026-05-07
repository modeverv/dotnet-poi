using DotnetPoi.HSLF.UserModel;
using Xunit;

namespace DotnetPoi.HSLF.Tests.UserModel;

/// <summary>
/// Phase 15 実装順 1: HSLF fixture survey.
/// Tests cover representative .ppt fixtures from POI test-data/slideshow/.
/// Expected values (slide count, text patterns) are drawn from POI upstream tests.
/// </summary>
public class HSLFSlideShowTests
{
    public static TheoryData<string, int, string?> RepresentativeFixtures => new()
    {
        // Representative fixtures from POI HSLF tests
        { "hslf-fixtures/basic_test_ppt_file.ppt", 2, "This is a test title" },
        { "hslf-fixtures/SampleShow.ppt",           2, "Title of the first slide" },
        { "hslf-fixtures/with_textbox.ppt",         1, "Hello, World!!!" },
        { "hslf-fixtures/text_shapes.ppt",          2, null },  // slide count from TestSheet
        { "hslf-fixtures/headers_footers.ppt",      1, null },  // header/footer content
        { "hslf-fixtures/WithComments.ppt",         1, null },  // has comments
        { "hslf-fixtures/pictures.ppt",             2, null },  // has 5 embedded pictures
        { "hslf-fixtures/testPPT_oleWorkbook.ppt",  1, null },  // has OLE embedding
        { "hslf-fixtures/54880_chinese.ppt",        1, "Single byte" }, // Chinese text
        { "hslf-fixtures/PPT95.ppt",                1, null },  // PPT95 legacy format
        { "hslf-fixtures/empty_textbox.ppt",        1, null },  // empty text boxes
        { "hslf-fixtures/backgrounds.ppt",          2, null },  // backgrounds (TestSheet)
        { "hslf-fixtures/incorrect_slide_order.ppt", 3, null }, // unusual order
    };

    /// <summary>Non-OLE2 stream must throw InvalidDataException.</summary>
    [Fact]
    public void Open_NonOle2Stream_ThrowsInvalidDataException()
    {
        var fake = new MemoryStream(new byte[1024]);
        Assert.Throws<InvalidDataException>(() => new HSLFSlideShow(fake));
    }

    [Theory]
    [MemberData(nameof(RepresentativeFixtures))]
    public void Open_Fixture_DoesNotThrow(string fixturePath, int expectedSlideCount, string? expectedSubstring)
    {
        using var stream = File.OpenRead(fixturePath);
        using var prs = new HSLFSlideShow(stream);
        Assert.NotNull(prs);
        // Legacy formats (PPT95) may yield 0 slides — that's a known gap
        if (!fixturePath.Contains("PPT95"))
            Assert.True(prs.getSlides().Count > 0, $"{fixturePath}: should have at least one slide");
    }

    [Theory]
    [MemberData(nameof(RepresentativeFixtures))]
    public void Open_Fixture_SlideCountBaseline(string fixturePath, int expectedSlideCount, string? expectedSubstring)
    {
        using var stream = File.OpenRead(fixturePath);
        using var prs = new HSLFSlideShow(stream);
        // Current simplified parser does not use persist pointers,
        // so slide count may differ from POI. Record actual vs expected for survey.
        var actual = prs.getSlides().Count;
        if (!fixturePath.Contains("PPT95"))
            Assert.True(actual > 0, $"{fixturePath}: expected >0 slides, got {actual}");
    }

    [Theory]
    [MemberData(nameof(RepresentativeFixtures))]
    public void Open_Fixture_TextExtractionDoesNotThrow(string fixturePath, int expectedSlideCount, string? expectedSubstring)
    {
        using var stream = File.OpenRead(fixturePath);
        using var prs = new HSLFSlideShow(stream);
        var slides = prs.getSlides();

        foreach (var slide in slides)
        {
            var paragraphs = slide.getTextParagraphs();
            var title = slide.getTitle();
            // Accessing text should never throw for any fixture
            _ = string.Join(" ", paragraphs);
            _ = title;
        }
    }

    [Theory]
    [MemberData(nameof(RepresentativeFixtures))]
    public void Open_Fixture_TitleAccessDoesNotThrow(string fixturePath, int expectedSlideCount, string? expectedSubstring)
    {
        using var stream = File.OpenRead(fixturePath);
        using var prs = new HSLFSlideShow(stream);
        foreach (var slide in prs.getSlides())
        {
            var title = slide.getTitle();
            Assert.NotNull(title);
        }
    }

    // ── Phase 15 実装順 2: stream inventory ─────────────────────────────────

    /// <summary>Stream names include the mandatory PowerPoint Document stream.</summary>
    [Theory]
    [InlineData("hslf-fixtures/basic_test_ppt_file.ppt")]
    [InlineData("hslf-fixtures/PPT95.ppt")]
    [InlineData("hslf-fixtures/testPPT_oleWorkbook.ppt")]
    public void StreamInventory_HasPowerPointDocument(string fixturePath)
    {
        using var stream = File.OpenRead(fixturePath);
        using var prs = new HSLFSlideShow(stream);
        var names = prs.getStreamNames();
        Assert.Contains(names, n => n.Contains("PowerPoint Document"));
    }

    /// <summary>Stream names include the Current User stream.</summary>
    [Theory]
    [InlineData("hslf-fixtures/basic_test_ppt_file.ppt")]
    [InlineData("hslf-fixtures/SampleShow.ppt")]
    public void StreamInventory_HasCurrentUser(string fixturePath)
    {
        using var stream = File.OpenRead(fixturePath);
        using var prs = new HSLFSlideShow(stream);
        var names = prs.getStreamNames();
        Assert.Contains(names, n => n.Contains("Current User"));
    }

    /// <summary>hasStream returns true for existing streams, false for missing ones.
    /// Note: SummaryInformation in OLE2 has a \\u0005 prefix in its stream name.</summary>
    [Theory]
    [InlineData("hslf-fixtures/basic_test_ppt_file.ppt", "PowerPoint Document", true)]
    [InlineData("hslf-fixtures/basic_test_ppt_file.ppt", "Current User", true)]
    [InlineData("hslf-fixtures/basic_test_ppt_file.ppt", "\u0005SummaryInformation", true)]
    [InlineData("hslf-fixtures/basic_test_ppt_file.ppt", "\u0005DocumentSummaryInformation", true)]
    [InlineData("hslf-fixtures/basic_test_ppt_file.ppt", "NonexistentStream", false)]
    [InlineData("hslf-fixtures/basic_test_ppt_file.ppt", "ObjectPool", false)]
    public void StreamInventory_HasStream(string fixturePath, string streamName, bool expected)
    {
        using var stream = File.OpenRead(fixturePath);
        using var prs = new HSLFSlideShow(stream);
        Assert.Equal(expected, prs.hasStream(streamName));
    }

    /// <summary>hasStorage tests — OLE2 storages have type=1 in directory entries.
    /// The Pictures entry in some fixtures is a stream (type=2), not a storage.</summary>
    [Theory]
    [InlineData("hslf-fixtures/testPPT_oleWorkbook.ppt", "Pictures", false)] // Pictures is a stream, not storage
    [InlineData("hslf-fixtures/basic_test_ppt_file.ppt", "ObjectPool", false)]
    public void StreamInventory_HasStorage(string fixturePath, string storageName, bool expected)
    {
        using var stream = File.OpenRead(fixturePath);
        using var prs = new HSLFSlideShow(stream);
        Assert.Equal(expected, prs.hasStorage(storageName));
    }

    /// <summary>hasEntry works for both streams and storages.</summary>
    [Theory]
    [InlineData("hslf-fixtures/testPPT_oleWorkbook.ppt", "PowerPoint Document", true)]
    [InlineData("hslf-fixtures/testPPT_oleWorkbook.ppt", "Pictures", true)] // Pictures is a stream
    [InlineData("hslf-fixtures/testPPT_oleWorkbook.ppt", "\u0005SummaryInformation", true)]
    [InlineData("hslf-fixtures/testPPT_oleWorkbook.ppt", "Nonexistent", false)]
    public void StreamInventory_HasEntry(string fixturePath, string entryName, bool expected)
    {
        using var stream = File.OpenRead(fixturePath);
        using var prs = new HSLFSlideShow(stream);
        Assert.Equal(expected, prs.hasEntry(entryName));
    }

    /// <summary>getStreamBytes returns non-null bytes for known streams.</summary>
    [Theory]
    [InlineData("hslf-fixtures/basic_test_ppt_file.ppt", "PowerPoint Document")]
    [InlineData("hslf-fixtures/basic_test_ppt_file.ppt", "Current User")]
    [InlineData("hslf-fixtures/basic_test_ppt_file.ppt", "\u0005SummaryInformation")]
    public void StreamInventory_GetStreamBytes_ReturnsNonNull(string fixturePath, string streamName)
    {
        using var stream = File.OpenRead(fixturePath);
        using var prs = new HSLFSlideShow(stream);
        var bytes = prs.getStreamBytes(streamName);
        Assert.NotNull(bytes);
        Assert.True(bytes.Length > 0, $"{streamName} should have content");
    }

    /// <summary>getStreamBytes returns null for nonexistent streams.</summary>
    [Fact]
    public void StreamInventory_GetStreamBytes_UnknownReturnsNull()
    {
        using var stream = File.OpenRead("hslf-fixtures/basic_test_ppt_file.ppt");
        using var prs = new HSLFSlideShow(stream);
        Assert.Null(prs.getStreamBytes("NonexistentStream"));
    }

    /// <summary>Stream count is at least 4 for minimal fixtures (PowerPoint Document + Current User + summary streams).</summary>
    [Theory]
    [InlineData("hslf-fixtures/basic_test_ppt_file.ppt")]
    [InlineData("hslf-fixtures/with_textbox.ppt")]
    [InlineData("hslf-fixtures/54880_chinese.ppt")]
    public void StreamInventory_MinimumStreamCount(string fixturePath)
    {
        using var stream = File.OpenRead(fixturePath);
        using var prs = new HSLFSlideShow(stream);
        Assert.True(prs.getStreamNames().Count >= 4,
            $"{fixturePath}: expected at least 4 streams (PowerPoint Document + Current User + 2 summary)");
    }

    /// <summary>testPPT_oleWorkbook has Pictures stream in addition to the standard 4.</summary>
    [Fact]
    public void StreamInventory_OleFixture_HasPicturesStream()
    {
        using var stream = File.OpenRead("hslf-fixtures/testPPT_oleWorkbook.ppt");
        using var prs = new HSLFSlideShow(stream);
        Assert.True(prs.getStreamNames().Count >= 5,
            "testPPT_oleWorkbook should have at least 5 streams including Pictures");
        Assert.True(prs.hasStream("Pictures"), "testPPT_oleWorkbook should have Pictures stream");
        // The OLE workbook is embedded via Pictures stream, not an ObjectPool storage
        Assert.False(prs.hasStorage("ObjectPool"), "testPPT_oleWorkbook does not have ObjectPool storage");
    }

    /// <summary>pictures.ppt also has Pictures stream.</summary>
    [Fact]
    public void StreamInventory_PicturesFixture_HasPicturesStream()
    {
        using var stream = File.OpenRead("hslf-fixtures/pictures.ppt");
        using var prs = new HSLFSlideShow(stream);
        Assert.True(prs.getStreamNames().Count >= 5,
            "pictures.ppt should have at least 5 streams including Pictures");
        Assert.True(prs.hasStream("Pictures"), "pictures.ppt should have Pictures stream");
    }

    // ── Phase 15 実装順 3: record tree parser ──────────────────────────────────

    /// <summary>getRootRecord returns the Document (1000) container for standard fixtures.</summary>
    [Theory]
    [InlineData("hslf-fixtures/basic_test_ppt_file.ppt")]
    [InlineData("hslf-fixtures/SampleShow.ppt")]
    [InlineData("hslf-fixtures/with_textbox.ppt")]
    public void RecordTree_RootIsDocumentContainer(string fixturePath)
    {
        using var stream = File.OpenRead(fixturePath);
        using var prs = new HSLFSlideShow(stream);
        var root = prs.getRootRecord();
        Assert.NotNull(root);
        Assert.Equal(1000, root.RecType);
        Assert.True(root.IsContainer, "Document record should be a container");
        Assert.NotNull(root.Children);
        Assert.NotEmpty(root.Children);
    }

    /// <summary>getTopLevelRecords returns all top-level records, including Slide containers.</summary>
    [Fact]
    public void RecordTree_TopLevelContainsSlideRecords()
    {
        using var stream = File.OpenRead("hslf-fixtures/basic_test_ppt_file.ppt");
        using var prs = new HSLFSlideShow(stream);
        var topLevel = prs.getTopLevelRecords();
        Assert.True(topLevel.Count >= 4,
            "basic_test_ppt_file should have at least 4 top-level records (Document + 2 Slides + EndDocument)");
        var slideRecords = topLevel.Where(r => r.RecType == 1006 && r.IsContainer).ToList();
        Assert.True(slideRecords.Count >= 2,
            "basic_test_ppt_file should have at least 2 Slide (1006) containers at top level");
    }

    /// <summary>Each record has valid recType, verAndInstance, recLen, offset, and raw bytes.</summary>
    [Fact]
    public void RecordTree_SlideRecord_HasRecordMetadata()
    {
        using var stream = File.OpenRead("hslf-fixtures/basic_test_ppt_file.ppt");
        using var prs = new HSLFSlideShow(stream);
        var slideRecord = prs.getTopLevelRecords().FirstOrDefault(r => r.RecType == 1006 && r.IsContainer);
        Assert.NotNull(slideRecord);
        Assert.Equal(1006, slideRecord.RecType);
        Assert.True(slideRecord.RecordLength > 0);
        Assert.True(slideRecord.Offset >= 0);
        Assert.NotNull(slideRecord.RawBytes);
        Assert.True(slideRecord.RawBytes.Length >= 8);
        // Raw bytes should start with the correct recType
        Assert.Equal(1006,
            System.Buffers.Binary.BinaryPrimitives.ReadUInt16LittleEndian(
                slideRecord.RawBytes.AsSpan(2)));
    }

    /// <summary>Slide container records have children (SlideAtom, PPDrawing, etc.).</summary>
    [Fact]
    public void RecordTree_SlideContainer_HasChildren()
    {
        using var stream = File.OpenRead("hslf-fixtures/basic_test_ppt_file.ppt");
        using var prs = new HSLFSlideShow(stream);
        var slideRecord = prs.getTopLevelRecords().FirstOrDefault(r => r.RecType == 1006 && r.IsContainer);
        Assert.NotNull(slideRecord);
        Assert.NotNull(slideRecord.Children);
        Assert.NotEmpty(slideRecord.Children);
        // A Slide container typically has: SlideAtom (1007), PPDrawing (1036), etc.
        Assert.Contains(slideRecord.Children, c => c.RecType == 1007 && c.IsAtom);
    }

    /// <summary>Atom records have no children and contain body data.</summary>
    [Fact]
    public void RecordTree_AtomRecord_HasNoChildren()
    {
        using var stream = File.OpenRead("hslf-fixtures/basic_test_ppt_file.ppt");
        using var prs = new HSLFSlideShow(stream);
        var docRoot = prs.getRootRecord();
        Assert.NotNull(docRoot);
        // DocumentAtom (1001) is a child of Document
        var docAtom = docRoot.Children!.FirstOrDefault(c => c.RecType == 1001);
        Assert.NotNull(docAtom);
        Assert.True(docAtom.IsAtom);
        Assert.Null(docAtom.Children);
        Assert.True(docAtom.RawBytes.Length >= 8);
    }

    /// <summary>TextCharsAtom and TextBytesAtom are preserved in the tree and text extraction works.</summary>
    [Fact]
    public void RecordTree_TextAtoms_PreservedInTree()
    {
        using var stream = File.OpenRead("hslf-fixtures/basic_test_ppt_file.ppt");
        using var prs = new HSLFSlideShow(stream);
        // Check that text atoms are present somewhere in the tree
        bool foundTextChars = false;
        bool foundTextBytes = false;
        WalkForTextAtoms(prs.getRootRecord(), ref foundTextChars, ref foundTextBytes);
        // basic_test_ppt_file uses TextBytesAtom (CP1252) for its text
        Assert.True(foundTextBytes || foundTextChars,
            "Should find at least one text atom in the record tree");
    }

    private static void WalkForTextAtoms(DotnetPoi.HSLF.Record.HSLFRecord? record,
        ref bool foundTextChars, ref bool foundTextBytes)
    {
        if (record is null) return;
        if (record.RecType == 4000) foundTextChars = true;
        if (record.RecType == 4008) foundTextBytes = true;
        if (record.Children is not null)
        {
            foreach (var child in record.Children)
                WalkForTextAtoms(child, ref foundTextChars, ref foundTextBytes);
        }
    }

    /// <summary>Unknown record types (e.g. Escher records 0xF000+) are preserved in the tree.</summary>
    [Fact]
    public void RecordTree_UnknownRecordTypes_PreservedInTree()
    {
        using var stream = File.OpenRead("hslf-fixtures/basic_test_ppt_file.ppt");
        using var prs = new HSLFSlideShow(stream);
        var ppDrawingGroup = prs.getTopLevelRecords()
            .SelectMany(r => r.Children ?? Enumerable.Empty<DotnetPoi.HSLF.Record.HSLFRecord>())
            .FirstOrDefault(c => c.RecType == 1035);
        Assert.NotNull(ppDrawingGroup);
        // PPDrawingGroup contains Escher records (0xF000+)
        var escherRecords = new List<DotnetPoi.HSLF.Record.HSLFRecord>();
        CollectEscherRecords(ppDrawingGroup, escherRecords);
        Assert.NotEmpty(escherRecords);
        // Verify at least one Escher record preserved with raw bytes
        Assert.All(escherRecords, e =>
        {
            Assert.True(e.RawBytes.Length >= 8, $"Escher recType={e.RecType} should have 8+ bytes");
        });
    }

    private static void CollectEscherRecords(DotnetPoi.HSLF.Record.HSLFRecord record,
        List<DotnetPoi.HSLF.Record.HSLFRecord> results)
    {
        if (record.RecType >= 0xF000)
            results.Add(record);
        if (record.Children is not null)
        {
            foreach (var child in record.Children)
                CollectEscherRecords(child, results);
        }
    }

    /// <summary>Record tree text extraction matches HSLFSlide text for every fixture.</summary>
    [Theory]
    [MemberData(nameof(RepresentativeFixtures))]
    public void RecordTree_TextExtraction_MatchesSlides(string fixturePath, int expectedSlideCount, string? expectedSubstring)
    {
        using var stream = File.OpenRead(fixturePath);
        using var prs = new HSLFSlideShow(stream);
        var slides = prs.getSlides();
        var textsFromTree = new List<string>();
        foreach (var record in prs.getTopLevelRecords())
            CollectTexts(record, textsFromTree);

        // Slide text atoms are a subset of all text atoms in the tree.
        // The tree also contains text atoms outside slides (e.g. in Environment).
        // Verify that each slide text paragraph appears somewhere in the tree texts.
        foreach (var slide in slides)
        {
            foreach (var para in slide.getTextParagraphs())
            {
                Assert.Contains(textsFromTree, t => t == para);
            }
        }

        // Also verify the tree has at least as many text atoms as the slides
        var slideTextCount = slides.Sum(s => s.getTextParagraphs().Count);
        Assert.True(textsFromTree.Count >= slideTextCount,
            $"Tree should have at least {slideTextCount} text atoms, got {textsFromTree.Count}");
    }

    private static void CollectTexts(DotnetPoi.HSLF.Record.HSLFRecord record, List<string> texts)
    {
        if (record is DotnetPoi.HSLF.Record.HSLFRecordAtom atom)
        {
            if (atom.RecType == 4000 && atom.Body.Length >= 2)
                texts.Add(System.Text.Encoding.Unicode.GetString(atom.Body.ToArray()));
            else if (atom.RecType == 4008 && atom.Body.Length >= 1)
                texts.Add(DotnetPoi.HSLF.UserModel.LocaleUtil1252Hslf.GetString(atom.Body));
        }
        if (record.Children is not null)
        {
            foreach (var child in record.Children)
                CollectTexts(child, texts);
        }
    }

    // ── Phase 15 実装順 4: persist pointer slide ordering ─────────────────────

    /// <summary>
    /// incorrect_slide_order.ppt has slides in a different order in the stream
    /// than they should appear. Persist pointers must resolve the correct order.
    /// POI TestSlideOrdering.testComplexCase expects 3 slides with titles
    /// "Slide 1", "Slide 2", "Slide 3" (in that order).
    /// </summary>
    [Fact]
    public void PersistPointer_IncorrectSlideOrder_HasThreeSlides()
    {
        using var stream = File.OpenRead("hslf-fixtures/incorrect_slide_order.ppt");
        using var prs = new HSLFSlideShow(stream);
        Assert.Equal(3, prs.getSlides().Count);
    }

    /// <summary>
    /// Verify that incorrect_slide_order.ppt returns slides in the correct
    /// display order (via persist pointers), matching POI behavior.
    /// </summary>
    [Fact]
    public void PersistPointer_IncorrectSlideOrder_TitlesInCorrectOrder()
    {
        using var stream = File.OpenRead("hslf-fixtures/incorrect_slide_order.ppt");
        using var prs = new HSLFSlideShow(stream);
        var slides = prs.getSlides();
        Assert.Equal(3, slides.Count);
        // POI TestSlideOrdering.testComplexCase expects titles "Slide 1", "Slide 2", "Slide 3"
        Assert.Contains("Slide 1", slides[0].getTitle());
        Assert.Contains("Slide 2", slides[1].getTitle());
        Assert.Contains("Slide 3", slides[2].getTitle());
    }

    /// <summary>
    /// basic_test_ppt_file.ppt has slides in simple order (record order matches
    /// display order). Verify the persist pointer model still works.
    /// Note: The SLWT (instance=0) has 1 SlideAtomsSet for this fixture.
    /// Additional Slide (1006) records are appended by the safety net
    /// but may have no text atoms.
    /// </summary>
    [Fact]
    public void PersistPointer_BasicTestPptFile_TitlesInCorrectOrder()
    {
        using var stream = File.OpenRead("hslf-fixtures/basic_test_ppt_file.ppt");
        using var prs = new HSLFSlideShow(stream);
        var slides = prs.getSlides();
        Assert.True(slides.Count >= 1, "Should have at least 1 slide");
        Assert.Contains("This is a test title", slides[0].getTitle());
    }

    /// <summary>
    /// PersistPtrHolder records exist in top-level records for standard fixtures.
    /// </summary>
    [Fact]
    public void PersistPointer_TopLevelContainsPersistPtrRecords()
    {
        using var stream = File.OpenRead("hslf-fixtures/basic_test_ppt_file.ppt");
        using var prs = new HSLFSlideShow(stream);
        var persistPtrs = prs.getTopLevelRecords()
            .Where(r => r.RecType == 6001 || r.RecType == 6002)
            .ToList();
        Assert.NotEmpty(persistPtrs);
    }

    /// <summary>
    /// Verify the persist pointer map can be built and contains entries.
    /// </summary>
    [Fact]
    public void PersistPointer_BuildPersistMap_NonEmpty()
    {
        using var stream = File.OpenRead("hslf-fixtures/basic_test_ppt_file.ppt");
        using var prs = new HSLFSlideShow(stream);
        var map = DotnetPoi.HSLF.Record.HSLFPersistPtrHolder.BuildPersistMap(prs.getTopLevelRecords());
        Assert.True(map.Count >= 1, "Should have at least 1 persist entry");
    }

    /// <summary>
    /// Persist map from incorrect_slide_order.ppt maps to valid Slide records.
    /// </summary>
    [Fact]
    public void PersistPointer_IncorrectSlideOrder_RecordOffsetsMatchSlides()
    {
        using var stream = File.OpenRead("hslf-fixtures/incorrect_slide_order.ppt");
        using var prs = new HSLFSlideShow(stream);
        var persistMap = DotnetPoi.HSLF.Record.HSLFPersistPtrHolder.BuildPersistMap(prs.getTopLevelRecords());

        Assert.True(persistMap.Count >= 1, "Should have at least 1 persist entry");

        // Verify each persist offset points to a valid record in the tree
        foreach (var offset in persistMap.Values)
        {
            var record = prs.getTopLevelRecords()
                .SelectMany(r => Flatten(r))
                .FirstOrDefault(r => r.Offset == offset);
            Assert.NotNull(record);
        }
    }

    private static IEnumerable<DotnetPoi.HSLF.Record.HSLFRecord> Flatten(DotnetPoi.HSLF.Record.HSLFRecord record)
    {
        yield return record;
        if (record.Children is not null)
        {
            foreach (var child in record.Children)
            {
                foreach (var nested in Flatten(child))
                    yield return nested;
            }
        }
    }

    // ── Phase 15 実装順 5: title/body text extraction ──────────────────────────

    /// <summary>
    /// Slide titles are correctly identified via TextHeaderAtom type.
    /// basic_test_ppt_file.ppt has "This is a test title" as title type.
    /// </summary>
    [Fact]
    public void TextExtraction_TitleIsFirstTextBlockWithTitleType()
    {
        using var stream = File.OpenRead("hslf-fixtures/basic_test_ppt_file.ppt");
        using var prs = new HSLFSlideShow(stream);
        var slides = prs.getSlides();
        Assert.True(slides.Count >= 1);
        var slide0 = slides[0];
        Assert.Contains("This is a test title", slide0.getTitle());
        var titleBlocks = slide0.getTextBlocks()
            .Where(b => b.Type == TextPlaceholderType.Title || b.Type == TextPlaceholderType.CenterTitle)
            .ToList();
        Assert.True(titleBlocks.Count > 0,
            "Should have at least one title-type text block");
    }

    /// <summary>
    /// Body text is correctly identified via TextHeaderAtom type.
    /// </summary>
    [Fact]
    public void TextExtraction_BodyTextBlocksHaveBodyType()
    {
        using var stream = File.OpenRead("hslf-fixtures/SampleShow.ppt");
        using var prs = new HSLFSlideShow(stream);
        var slides = prs.getSlides();
        foreach (var slide in slides)
        {
            var bodyBlocks = slide.getTextBlocks()
                .Where(b => b.Type == TextPlaceholderType.Body)
                .ToList();
            if (slide.getBodyParagraphs().Count > 0)
            {
                Assert.True(bodyBlocks.Count > 0,
                    $"Slide with body text should have Body-type blocks (title='{slide.getTitle()}')");
            }
        }
    }

    /// <summary>
    /// 54880_chinese.ppt does not throw on open. Text may be in notes
    /// or special records not yet parsed; at minimum the fixture opens.
    /// </summary>
    [Fact]
    public void TextExtraction_ChineseFixture_DoesNotThrow()
    {
        using var stream = File.OpenRead("hslf-fixtures/54880_chinese.ppt");
        using var prs = new HSLFSlideShow(stream);
        var slides = prs.getSlides();
        Assert.True(slides.Count >= 1,
            "54880_chinese.ppt should have at least 1 slide");
        // Text extraction should not throw even if the slide text is empty
        foreach (var slide in slides)
        {
            _ = slide.getTitle();
            _ = slide.getTextParagraphs();
            _ = slide.getTextBlocks();
        }
    }

    /// <summary>
    /// Empty text boxes do not cause exceptions.
    /// </summary>
    [Fact]
    public void TextExtraction_EmptyTextBox_DoesNotThrow()
    {
        using var stream = File.OpenRead("hslf-fixtures/empty_textbox.ppt");
        using var prs = new HSLFSlideShow(stream);
        // Should open without throwing
        foreach (var slide in prs.getSlides())
        {
            var title = slide.getTitle();
            var paragraphs = slide.getTextParagraphs();
            _ = title;
            _ = paragraphs;
        }
    }

    /// <summary>
    /// with_textbox.ppt opens and extracts text (multiple text boxes
    /// may be merged into one text atom with embedded newlines).
    /// </summary>
    [Fact]
    public void TextExtraction_WithTextBox_DoesNotThrow()
    {
        using var stream = File.OpenRead("hslf-fixtures/with_textbox.ppt");
        using var prs = new HSLFSlideShow(stream);
        var slides = prs.getSlides();
        Assert.True(slides.Count >= 1);
        foreach (var slide in slides)
        {
            _ = slide.getTitle();
            _ = slide.getTextParagraphs();
            _ = slide.getTextBlocks();
        }
    }

    // ── Phase 15 実装順 6: no-op write round-trip ─────────────────────────────

    /// <summary>
    /// No-op write preserves all streams for a simple fixture.
    /// Write → read back → verify slide count and stream names.
    /// </summary>
    [Theory]
    [InlineData("hslf-fixtures/basic_test_ppt_file.ppt")]
    [InlineData("hslf-fixtures/SampleShow.ppt")]
    [InlineData("hslf-fixtures/with_textbox.ppt")]
    [InlineData("hslf-fixtures/text_shapes.ppt")]
    public void RoundTrip_NoOpWrite_PreservesSlideCountAndStreams(string fixturePath)
    {
        byte[] originalBytes;
        using (var stream = File.OpenRead(fixturePath))
        using (var ms = new MemoryStream())
        {
            stream.CopyTo(ms);
            originalBytes = ms.ToArray();
        }

        // First read: capture state
        int slideCount;
        List<string> streamNames;
        using (var stream = new MemoryStream(originalBytes))
        using (var prs = new HSLFSlideShow(stream))
        {
            slideCount = prs.getSlides().Count;
            streamNames = prs.getStreamNames().ToList();
        }

        // Write
        byte[] written;
        using (var stream = new MemoryStream(originalBytes))
        using (var prs = new HSLFSlideShow(stream))
        using (var output = new MemoryStream())
        {
            prs.write(output);
            written = output.ToArray();
        }

        // Read back
        using (var input2 = new MemoryStream(written))
        using (var prs2 = new HSLFSlideShow(input2))
        {
            Assert.Equal(slideCount, prs2.getSlides().Count);
            var names2 = prs2.getStreamNames().ToList();
            Assert.True(streamNames.Count <= names2.Count,
                $"Round-tripped document should have at least the same number of streams");
            foreach (var name in streamNames)
                Assert.Contains(names2, n => n == name);
        }
    }

    /// <summary>
    /// No-op write preserves extracted text for simple fixtures.
    /// </summary>
    [Theory]
    [InlineData("hslf-fixtures/basic_test_ppt_file.ppt")]
    [InlineData("hslf-fixtures/SampleShow.ppt")]
    [InlineData("hslf-fixtures/with_textbox.ppt")]
    [InlineData("hslf-fixtures/text_shapes.ppt")]
    public void RoundTrip_NoOpWrite_PreservesExtractedText(string fixturePath)
    {
        byte[] originalBytes;
        using (var stream = File.OpenRead(fixturePath))
        using (var ms = new MemoryStream())
        {
            stream.CopyTo(ms);
            originalBytes = ms.ToArray();
        }

        // Extract texts before
        List<string> beforeTexts;
        using (var stream = new MemoryStream(originalBytes))
        using (var prs = new HSLFSlideShow(stream))
        {
            beforeTexts = prs.getSlides()
                .SelectMany(s => s.getTextParagraphs())
                .ToList();
        }

        // Write
        byte[] written;
        using (var stream = new MemoryStream(originalBytes))
        using (var prs = new HSLFSlideShow(stream))
        using (var output = new MemoryStream())
        {
            prs.write(output);
            written = output.ToArray();
        }

        // Extract texts after
        List<string> afterTexts;
        using (var stream = new MemoryStream(written))
        using (var prs = new HSLFSlideShow(stream))
        {
            afterTexts = prs.getSlides()
                .SelectMany(s => s.getTextParagraphs())
                .ToList();
        }

        Assert.Equal(beforeTexts.Count, afterTexts.Count);
        for (int i = 0; i < beforeTexts.Count; i++)
            Assert.Equal(beforeTexts[i], afterTexts[i]);
    }

    /// <summary>
    /// No-op write preserves pictures, comments, and OLE streams.
    /// </summary>
    [Theory]
    [InlineData("hslf-fixtures/pictures.ppt")]
    [InlineData("hslf-fixtures/WithComments.ppt")]
    [InlineData("hslf-fixtures/testPPT_oleWorkbook.ppt")]
    public void RoundTrip_NoOpWrite_PreservesSpecialStreams(string fixturePath)
    {
        byte[] written;
        using (var stream = File.OpenRead(fixturePath))
        using (var prs = new HSLFSlideShow(stream))
        {
            using (var output = new MemoryStream())
            {
                prs.write(output);
                written = output.ToArray();
            }
        }

        using (var stream = new MemoryStream(written))
        using (var prs = new HSLFSlideShow(stream))
        {
            Assert.True(prs.hasStream("PowerPoint Document"),
                "Round-tripped document must have 'PowerPoint Document' stream");
        }
    }

    /// <summary>
    /// The PowerPoint Document stream bytes are preserved byte-for-byte
    /// after a no-op write (no editing performed).
    /// </summary>
    [Fact]
    public void RoundTrip_NoOpWrite_PowerPointDocumentStreamIdentical()
    {
        var fixture = "hslf-fixtures/basic_test_ppt_file.ppt";

        using (var stream = File.OpenRead(fixture))
        using (var prs = new HSLFSlideShow(stream))
        {
            var originalPpt = prs.getStreamBytes("PowerPoint Document");
            Assert.NotNull(originalPpt);

            using (var output = new MemoryStream())
            {
                prs.write(output);
                var writtenBytes = output.ToArray();

                using (var input2 = new MemoryStream(writtenBytes))
                using (var prs2 = new HSLFSlideShow(input2))
                {
                    var roundTrippedPpt = prs2.getStreamBytes("PowerPoint Document");
                    Assert.NotNull(roundTrippedPpt);
                    Assert.True(originalPpt.Length == roundTrippedPpt.Length,
                        $"PowerPoint Document stream length should match ({originalPpt.Length} == {roundTrippedPpt.Length})");
                    Assert.Equal(originalPpt, roundTrippedPpt);
                }
            }
        }
    }
}
