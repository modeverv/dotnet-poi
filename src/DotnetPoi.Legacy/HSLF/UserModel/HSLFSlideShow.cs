using System.Buffers.Binary;
using System.Linq;
using System.Text;
using DotnetPoi.HSLF.Record;
using DotnetPoi.POIFS.Crypt;

namespace DotnetPoi.HSLF.UserModel;

/// <summary>
/// Minimal read-only port of org.apache.poi.hslf.usermodel.HSLFSlideShow.
/// Reads PowerPoint 97–2003 binary (.ppt) files and extracts text and slide count.
///
/// Record structure reference: [MS-PPT] § 2.1.1 (RecordHeader), § 2.4.14.1
/// (TextCharsAtom), § 2.4.14.2 (TextBytesAtom).
/// </summary>
public sealed class HSLFSlideShow : IDisposable
{
    private const string StreamPowerPointDocument = "PowerPoint Document";

    // Record type constants from RecordTypes.java
    internal const ushort RecTypeSlide             = 1006;
    internal const ushort RecTypeSlidePersistAtom  = 1011;
    internal const ushort RecTypeSlideListWithText = 4080;
    internal const ushort RecTypeTextHeaderAtom    = 3999;
    internal const ushort RecTypeTextCharsAtom     = 4000;  // UTF-16LE text
    internal const ushort RecTypeTextBytesAtom     = 4008;  // CP1252 text
    internal const byte   ContainerVersionFlag     = 0x0F;  // low nibble of verAndInstance

    private readonly CompoundFileDocument _fileSystem;
    private readonly IReadOnlyList<HSLFRecord> _topLevelRecords;
    private readonly byte[] _pptStream;
    private readonly List<HSLFSlide> _slides = new();

    public HSLFSlideShow(Stream stream)
    {
        Guard.ThrowIfNull(stream, nameof(stream));
        _fileSystem = CompoundFile.ReadDocument(stream);
        if (!_fileSystem.Streams.TryGetValue(StreamPowerPointDocument, out var pptStream))
            throw new InvalidDataException("Not a valid .ppt file: missing 'PowerPoint Document' stream.");

        _pptStream = pptStream;

        // Parse all top-level records in the PowerPoint Document stream
        var records = new List<HSLFRecord>();
        int pos = 0;
        while (pos + HSLFRecord.RawBytesHeaderSize <= pptStream.Length)
        {
            var record = ParseRecord(pptStream, ref pos);
            if (record is null) break;
            records.Add(record);
        }

        _topLevelRecords = records;

        // Phase 15 実装順 4: Build slides using persist pointers for correct order
        BuildSlidesWithPersistPointers();
    }

    /// <summary>
    /// Returns the Document (1000) root record, or null if not found.
    /// The Document record is typically the first top-level record.
    /// </summary>
    public HSLFRecord? getRootRecord() =>
        _topLevelRecords.FirstOrDefault(r => r.RecType == 1000);

    /// <summary>Returns all top-level records in the PowerPoint Document stream.</summary>
    public IReadOnlyList<HSLFRecord> getTopLevelRecords() => _topLevelRecords;

    /// <summary>Returns the slides in the presentation.</summary>
    public IReadOnlyList<HSLFSlide> getSlides() => _slides;

    /// <summary>Returns all OLE2 stream names (path-qualified) in the compound document.</summary>
    public IReadOnlyCollection<string> getStreamNames() =>
        _fileSystem.Streams.Keys.OrderBy(name => name, StringComparer.Ordinal).ToArray();

    /// <summary>Returns true when the OLE2 document has a stream with the exact path/name.</summary>
    public bool hasStream(string name) => _fileSystem.Streams.ContainsKey(name);

    /// <summary>Returns true when the OLE2 document has a storage with the exact path/name.</summary>
    public bool hasStorage(string name) =>
        _fileSystem.EntryMetadata.TryGetValue(name, out var metadata) && metadata.Type == 1;

    /// <summary>Returns true when the OLE2 document has a stream or storage with the exact path/name.</summary>
    public bool hasEntry(string name) => hasStream(name) || hasStorage(name);

    /// <summary>
    /// Returns the raw bytes of a named OLE2 stream, or null if not found.
    /// Path format uses '/' as separator (e.g. "PowerPoint Document", "ObjectPool/_1494912778").
    /// </summary>
    public byte[]? getStreamBytes(string name) =>
        _fileSystem.Streams.TryGetValue(name, out var data) ? data : null;

    /// <summary>
    /// No-op write: preserves all OLE2 streams byte-for-byte without modifying
    /// the PowerPoint Document stream.
    /// </summary>
    public void write(Stream stream)
    {
        Guard.ThrowIfNull(stream, nameof(stream));
        CompoundFile.Write(stream, _fileSystem);
    }

    public void Dispose() { }

    /// <summary>
    /// Builds the slide list using persist pointers for correct slide order.
    ///
    /// PowerPoint stores slide text in the Document container's SlideListWithText (4080),
    /// NOT inside the Slide (1006) record tree. The text blocks are grouped after each
    /// SlidePersistAtom (1011) in the SLWT. The SlidePersistAtom's refID maps to a persist
    /// ID that points to the Slide (1006) record offset.
    ///
    /// Algorithm:
    /// 1. Parse PersistPtrIncrementalBlock (6002) records → persistId → offset map
    /// 2. Find Document (1000) container
    /// 3. Find SlideListWithText (4080, instance=0 = slides) inside Document
    /// 4. Walk SLWT children, grouping into SlideAtomsSets (SlidePersistAtom + text atoms)
    /// 5. For each SlideAtomsSet in order:
    ///    - Read refID from SlidePersistAtom
    ///    - Extract text from TextBytesAtom/TextCharsAtom
    ///    - Read TextHeaderAtom type to classify title vs body vs notes
    ///    - Try to find matching Slide (1006) via persist map
    ///    - Create HSLFSlide with typed text blocks
    /// 6. Fall back to record-appearance order if the SLWT-based approach fails
    /// </summary>
    private void BuildSlidesWithPersistPointers()
    {
        // Step 1: Build persist ID → offset map
        var persistMap = HSLFPersistPtrHolder.BuildPersistMap(_topLevelRecords);

        // Step 2: Find Document container
        var docRecord = getRootRecord();
        if (docRecord is null)
            return;

        // Step 3: Find SlideListWithText (instance = 0 = slides) inside Document
        var slideSLWT = FindSlideListWithText(docRecord, instance: 0);
        if (slideSLWT is null || slideSLWT.Children is null || slideSLWT.Children.Count == 0)
        {
            // Fallback: use record-appearance order
            foreach (var record in _topLevelRecords)
                ExtractSlidesFallback(record);
            return;
        }

        // Step 4: Walk SLWT children, grouping into SlideAtomsSets.
        // Each set starts with a SlidePersistAtom (1011), followed by
        // text-related atoms (3999=TextHeaderAtom, 4008=TextBytesAtom,
        // 4000=TextCharsAtom, 4010=StyleTextPropAtom) until the next
        // SlidePersistAtom or end of children.

        var slideSets = new List<(int RefId, List<HSLFRecord> TextBlocks)>();
        HSLFRecord? currentSPA = null;
        var currentTexts = new List<HSLFRecord>();

        foreach (var child in slideSLWT.Children)
        {
            if (child.RecType == RecTypeSlidePersistAtom)
            {
                // Finalize previous slide set
                if (currentSPA is not null)
                {
                    int refId = currentSPA.RecordLength >= 4
                        ? BinaryPrimitives.ReadInt32LittleEndian(currentSPA.Body.Slice(0, 4))
                        : 0;
                    slideSets.Add((refId, currentTexts));
                    currentTexts = new List<HSLFRecord>();
                }
                currentSPA = child;
            }
            else if (currentSPA is not null)
            {
                currentTexts.Add(child);
            }
        }

        // Finalize the last slide
        if (currentSPA is not null)
        {
            int refId = currentSPA.RecordLength >= 4
                ? BinaryPrimitives.ReadInt32LittleEndian(currentSPA.Body.Slice(0, 4))
                : 0;
            slideSets.Add((refId, currentTexts));
        }

        if (slideSets.Count == 0)
        {
            // No SlidePersistAtoms found — fallback
            foreach (var record in _topLevelRecords)
                ExtractSlidesFallback(record);
            return;
        }

        // Step 5: Build slide lookup from persist offset → Slide (1006) record
        var slideByOffset = new Dictionary<int, HSLFRecord>();
        foreach (var record in _topLevelRecords)
            CollectSlidesByOffset(record, slideByOffset);

        var usedOffsets = new HashSet<int>();

        // Step 6: For each SlideAtomsSet in SLWT order, create HSLFSlide
        foreach (var (refId, textBlocks) in slideSets)
        {
            // Extract text blocks with their TextHeaderAtom type
            var slideBlocks = new List<HSLFTextBlock>();
            TextPlaceholderType currentType = TextPlaceholderType.Body; // default

            foreach (var block in textBlocks)
            {
                if (block is HSLFRecordAtom atom)
                {
                    if (atom.RecType == RecTypeTextHeaderAtom && atom.Body.Length >= 4)
                    {
                        // TextHeaderAtom stores the text type as a 4-byte int32
                        int rawType = BinaryPrimitives.ReadInt32LittleEndian(atom.Body.Slice(0, 4));
                        currentType = (TextPlaceholderType)rawType;
                    }
                    else if (atom.RecType == RecTypeTextBytesAtom && atom.Body.Length >= 1)
                    {
                        string text = LocaleUtil1252Hslf.GetString(atom.Body);
                        slideBlocks.Add(new HSLFTextBlock(text, currentType));
                        // After a text atom, reset type to Body (next text without explicit header
                        // is assumed body)
                        currentType = TextPlaceholderType.Body;
                    }
                    else if (atom.RecType == RecTypeTextCharsAtom && atom.Body.Length >= 2)
                    {
                        string text = Encoding.Unicode.GetString(atom.Body.ToArray());
                        slideBlocks.Add(new HSLFTextBlock(text, currentType));
                        currentType = TextPlaceholderType.Body;
                    }
                }
            }

            // Track the slide record if resolvable via persist pointer
            if (persistMap.TryGetValue(refId, out int offset) && slideByOffset.ContainsKey(offset))
            {
                usedOffsets.Add(offset);
            }

            _slides.Add(new HSLFSlide(slideBlocks));
        }

        // Append any Slide records not referenced by persist pointers
        // (safety net for edge cases)
        foreach (var kv in slideByOffset)
        {
            if (!usedOffsets.Contains(kv.Key))
            {
                var slideBlocks = new List<HSLFTextBlock>();
                ExtractBlocksFromTree(kv.Value, slideBlocks);
                _slides.Add(new HSLFSlide(slideBlocks));
            }
        }
    }

    /// <summary>
    /// Finds the SlideListWithText record inside the Document container matching the given instance.
    /// Instance values: 0 = slides, 1 = master slides, 2 = notes.
    /// The instance field is encoded in the high nibble of the verAndInstance header word.
    /// </summary>
    private static HSLFRecord? FindSlideListWithText(HSLFRecord docRecord, int instance)
    {
        if (docRecord.Children is null)
            return null;

        foreach (var child in docRecord.Children)
        {
            if (child.RecType != RecTypeSlideListWithText)
                continue;
            if (!child.IsContainer)
                continue;

            // Instance = verAndInstance >> 4
            int inst = child.VerAndInstance >> 4;
            if (inst == instance)
                return child;
        }

        return null;
    }

    /// <summary>
    /// Recursively collects Slide (1006) container records keyed by their byte offset.
    /// </summary>
    private static void CollectSlidesByOffset(HSLFRecord record, Dictionary<int, HSLFRecord> map)
    {
        if (record.IsContainer && record.RecType == RecTypeSlide)
            map[record.Offset] = record;

        if (record.Children is not null)
        {
            foreach (var child in record.Children)
                CollectSlidesByOffset(child, map);
        }
    }

    // ─── fallback (old-style, record-appearance order) ─────────────────────

    private void ExtractSlidesFallback(HSLFRecord record)
    {
        if (record.IsContainer && record.RecType == RecTypeSlide)
        {
            var slideBlocks = new List<HSLFTextBlock>();
            ExtractBlocksFromTree(record, slideBlocks);
            _slides.Add(new HSLFSlide(slideBlocks));
        }

        if (record.Children is not null)
        {
            foreach (var child in record.Children)
                ExtractSlidesFallback(child);
        }
    }

    // ─── record tree parsing ────────────────────────────────────────────────

    /// <summary>
    /// Parses a single record (atom or container) from the byte array at the given position.
    /// Advances the position past the record.
    /// </summary>
    private static HSLFRecord? ParseRecord(byte[] data, ref int pos)
    {
        int start = pos;
        if (pos + HSLFRecord.RawBytesHeaderSize > data.Length)
            return null;

        var verAndInst = BinaryPrimitives.ReadUInt16LittleEndian(data.AsSpan(pos));
        var recType    = BinaryPrimitives.ReadUInt16LittleEndian(data.AsSpan(pos + 2));
        var recLen     = BinaryPrimitives.ReadInt32LittleEndian(data.AsSpan(pos + 4));
        pos += HSLFRecord.RawBytesHeaderSize;

        if (recLen < 0 || pos + recLen > data.Length)
        {
            // Invalid record — roll back and stop
            pos = start;
            return null;
        }

        bool isContainer = (verAndInst & 0x0F) == ContainerVersionFlag;

        if (isContainer)
        {
            // Parse children from the body
            var children = new List<HSLFRecord>();
            int bodyEnd = pos + recLen;
            while (pos + HSLFRecord.RawBytesHeaderSize <= bodyEnd)
            {
                var child = ParseRecord(data, ref pos);
                if (child is null) break;
                children.Add(child);
            }
            // If we stopped early, move to the body end
            if (pos < bodyEnd)
                pos = bodyEnd;

            // Raw bytes for container = header + all children serialized
            var rawLen = HSLFRecord.RawBytesHeaderSize + recLen;
            var rawBytes = new byte[rawLen];
            Array.Copy(data, start, rawBytes, 0, Math.Min(rawLen, data.Length - start));

            return new HSLFRecordContainer(recType, verAndInst, recLen, start, rawBytes, children);
        }
        else
        {
            // Atom record — body is raw data
            var rawLen = HSLFRecord.RawBytesHeaderSize + recLen;
            var rawBytes = new byte[rawLen];
            Array.Copy(data, start, rawBytes, 0, Math.Min(rawLen, data.Length - start));
            pos += recLen;

            return new HSLFRecordAtom(recType, verAndInst, recLen, start, rawBytes);
        }
    }

    // ─── text extraction from tree ─────────────────────────────────────────

    private static void ExtractTextsFromTree(HSLFRecord record, List<string> texts)
    {
        if (record is HSLFRecordAtom atom)
        {
            switch (atom.RecType)
            {
                case RecTypeTextCharsAtom:
                    if (atom.Body.Length >= 2)
                        texts.Add(Encoding.Unicode.GetString(atom.Body.ToArray()));
                    break;

                case RecTypeTextBytesAtom:
                    if (atom.Body.Length >= 1)
                        texts.Add(LocaleUtil1252Hslf.GetString(atom.Body));
                    break;
            }
        }

        if (record.Children is not null)
        {
            foreach (var child in record.Children)
                ExtractTextsFromTree(child, texts);
        }
    }

    /// <summary>
    /// Extracts typed text blocks from a record tree (used for fallback slides).
    /// Without a TextHeaderAtom context, blocks are assumed Body type.
    /// </summary>
    private static void ExtractBlocksFromTree(HSLFRecord record, List<HSLFTextBlock> blocks)
    {
        if (record is HSLFRecordAtom atom)
        {
            switch (atom.RecType)
            {
                case RecTypeTextCharsAtom:
                    if (atom.Body.Length >= 2)
                        blocks.Add(new HSLFTextBlock(
                            Encoding.Unicode.GetString(atom.Body.ToArray()),
                            TextPlaceholderType.Body));
                    break;

                case RecTypeTextBytesAtom:
                    if (atom.Body.Length >= 1)
                        blocks.Add(new HSLFTextBlock(
                            LocaleUtil1252Hslf.GetString(atom.Body),
                            TextPlaceholderType.Body));
                    break;
            }
        }

        if (record.Children is not null)
        {
            foreach (var child in record.Children)
                ExtractBlocksFromTree(child, blocks);
        }
    }
}

/// <summary>
/// Identifies the type of a text block on a slide, matching POI's
/// TextHeaderAtom.rtCharsType / TextPlaceholder constants.
/// </summary>
public enum TextPlaceholderType
{
    Title       = 0,
    Body        = 1,
    Notes       = 2,
    Other       = 3,
    CenterBody  = 5,
    CenterTitle = 6,
    HalfBody    = 7,
    QuarterBody = 8,
}

/// <summary>
/// Represents a single text block on a slide, with its type from TextHeaderAtom.
/// Ported from org.apache.poi.hslf.record.TextHeaderAtom.
/// </summary>
public sealed class HSLFTextBlock
{
    internal HSLFTextBlock(string text, TextPlaceholderType type)
    {
        Text = text ?? string.Empty;
        Type = type;
    }

    /// <summary>The decoded text content.</summary>
    public string Text { get; }

    /// <summary>The placeholder type (title, body, notes, etc.).</summary>
    public TextPlaceholderType Type { get; }
}

/// <summary>
/// Represents a single slide in a PowerPoint 97–2003 presentation.
/// Ported from org.apache.poi.hslf.usermodel.HSLFSlide.
/// </summary>
public sealed class HSLFSlide
{
    private readonly IReadOnlyList<HSLFTextBlock> _textBlocks;

    internal HSLFSlide(IReadOnlyList<HSLFTextBlock> textBlocks)
    {
        _textBlocks = textBlocks;
    }

    /// <summary>Returns all text runs on this slide as plain strings.</summary>
    public IReadOnlyList<string> getTextParagraphs() =>
        _textBlocks.Select(b => b.Text).ToList().AsReadOnly();

    /// <summary>Returns all text blocks with their placeholder type.</summary>
    public IReadOnlyList<HSLFTextBlock> getTextBlocks() => _textBlocks;

    /// <summary>
    /// Returns the title text (first TextPlaceholderType.Title or CenterTitle block),
    /// or empty string if no title block is found.
    /// </summary>
    public string getTitle()
    {
        var title = _textBlocks.FirstOrDefault(
            b => b.Type == TextPlaceholderType.Title ||
                 b.Type == TextPlaceholderType.CenterTitle);
        return title?.Text ?? string.Empty;
    }

    /// <summary>Returns body text paragraphs (Body, CenterBody, HalfBody, QuarterBody types).</summary>
    public IReadOnlyList<string> getBodyParagraphs() =>
        _textBlocks.Where(b => b.Type == TextPlaceholderType.Body ||
                               b.Type == TextPlaceholderType.CenterBody ||
                               b.Type == TextPlaceholderType.HalfBody ||
                               b.Type == TextPlaceholderType.QuarterBody)
                   .Select(b => b.Text).ToList().AsReadOnly();
}

internal static class LocaleUtil1252Hslf
{
    private static Encoding? _enc;

    internal static string GetString(ReadOnlySpan<byte> data)
    {
        if (_enc is null)
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            _enc = Encoding.GetEncoding(1252);
        }
        return _enc.GetString(data.ToArray());
    }
}
