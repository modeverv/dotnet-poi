using System.Buffers.Binary;
using System.Text;
using DotnetPoi.POIFS.Crypt;

namespace DotnetPoi.HWPF.UserModel;

/// <summary>
/// Minimal read-only port of org.apache.poi.hwpf.HWPFDocument.
/// Reads Word 97–2003 binary (.doc) files and extracts the main document text.
///
/// Algorithm reference: MS-DOC § 2.4 (FIB), § 2.8 (Text), § 2.9.196 (FcCompressed),
/// § 2.9.178 (Pcdt), § 2.9.176 (PieceDescriptor).
/// </summary>
public sealed class HWPFDocument : IDisposable
{
    // Stream names inside the OLE2 container
    private const string StreamWordDocument = "WordDocument";
    private const string Stream0Table = "0Table";
    private const string Stream1Table = "1Table";

    // Offsets within the WordDocument stream (fibBase + fibRgW97 + fibRgLw97)
    // fibBase = 32 bytes; fibRgW97 = 28 bytes; fibRgLw97 = 88 bytes
    private const int FibOffsetFlags1          = 10;   // 2-byte flags: bit9=fWhichTblStm
    private const int FibOffsetCcpText         = 72;   // fibRgLw97 base (60) + 0xC = 72
    private const int FibOffsetFcClx           = 418;  // fibRgFcLcb start (154) + CLX index 33 * 8
    private const int FibOffsetLcbClx          = 422;  // fcClx + 4
    // fibRgFcLcb97: starts at byte 154. Each entry = 8 bytes (fc:int32 + lcb:int32)
    // FIBFieldHandler constants: PLCFBTECHPX=12, STTBFFFN=15
    private const int FibOffsetFcPlcfBteChpx   = 154 + 12 * 8;      // 250
    private const int FibOffsetLcbPlcfBteChpx  = 154 + 12 * 8 + 4;  // 254
    private const int FibOffsetFcSttbfffn       = 154 + 15 * 8;      // 274
    private const int FibOffsetLcbSttbfffn      = 154 + 15 * 8 + 4;  // 278

    private byte[] _mainStream;
    private readonly CompoundFileDocument _fileSystem;
    private HWPFFileInformationBlock _fib;
    private HWPFTextModel _textModel;
    private string _text;
    private string[] _fontTable = Array.Empty<string>();
    private List<(int CpStart, int CpEnd, HWPFChpProperties Chp)> _chpSegments = new();

    public HWPFDocument(Stream stream)
    {
        Guard.ThrowIfNull(stream, nameof(stream));
        _fileSystem = CompoundFile.ReadDocument(stream);
        var streams = _fileSystem.Streams;
        if (!streams.TryGetValue(StreamWordDocument, out var mainStream))
            throw new InvalidDataException("Not a valid .doc file: missing WordDocument stream.");
        _mainStream = mainStream;
        _fib = HWPFFileInformationBlock.Read(mainStream, streams);
        _textModel = ExtractTextModel(mainStream, _fib);
        _text = _textModel.Text;
        _fontTable = ReadFontTable(mainStream, streams, _fib);
        _chpSegments = ReadChpSegments(mainStream, _fib, _textModel, _fontTable);
    }

    /// <summary>
    /// Returns the plain text of the main document body.
    /// Ported from HWPFDocument.getText().
    /// </summary>
    public string getText() => _text;

    /// <summary>Returns the character count of the main document body.</summary>
    public int getCcpText() => _fib.CcpText;

    /// <summary>Returns the main document range. Ported from HWPFDocument.getRange().</summary>
    public Range getRange() => new(0, _text.Length, _textModel, _chpSegments);

    /// <summary>Returns parsed File Information Block fields used by the current HWPF reader.</summary>
    public HWPFFileInformationBlock getFileInformationBlock() => _fib;

    /// <summary>Returns all OLE2 stream paths visible to the HWPF document.</summary>
    public IReadOnlyCollection<string> getStreamNames() => _fileSystem.Streams.Keys.OrderBy(name => name, StringComparer.Ordinal).ToArray();

    /// <summary>Returns true when the OLE2 document has a stream with the exact path/name.</summary>
    public bool hasStream(string name) => _fileSystem.Streams.ContainsKey(name);

    /// <summary>Returns true when the OLE2 document has a storage with the exact path/name.</summary>
    public bool hasStorage(string name) =>
        _fileSystem.EntryMetadata.TryGetValue(name, out var metadata) && metadata.Type == 1;

    /// <summary>Returns true when the OLE2 document has a stream or storage with the exact path/name.</summary>
    public bool hasEntry(string name) => hasStream(name) || hasStorage(name);

    /// <summary>
    /// Writes the loaded .doc back without rebuilding Word binary tables.
    /// This is the HWPF no-op preservation path; editing APIs are ported later.
    /// </summary>
    public void write(Stream stream)
    {
        Guard.ThrowIfNull(stream, nameof(stream));
        CompoundFile.Write(stream, _fileSystem);
    }

    /// <summary>
    /// Appends a plain paragraph to the main document body.
    /// Limited Phase 13 edit path: formatting tables are preserved but not rebuilt.
    /// </summary>
    public void appendParagraph(string text)
    {
        Guard.ThrowIfNull(text, nameof(text));
        var paragraphText = text.EndsWith("\r", StringComparison.Ordinal) ? text : text + "\r";
        SetMainBodyText(_text + paragraphText);
    }

    /// <summary>
    /// Replaces all occurrences of a plain-text placeholder in the main document body.
    /// Ported API shape from Range.replaceText(String, String), implemented at document scope for now.
    /// </summary>
    public void replaceText(string placeholder, string value)
    {
        Guard.ThrowIfNull(placeholder, nameof(placeholder));
        Guard.ThrowIfNull(value, nameof(value));
        if (placeholder.Length == 0) throw new ArgumentException("Placeholder must not be empty.", nameof(placeholder));
        if (_text.IndexOf(placeholder, StringComparison.Ordinal) < 0) return;

        SetMainBodyText(ReplaceOrdinal(_text, placeholder, value));
    }

    public void Dispose() { }

    // ─── text extraction ────────────────────────────────────────────────────

    private static HWPFTextModel ExtractTextModel(byte[] main, HWPFFileInformationBlock fib)
    {
        if (fib.SelectedTableStream is null || !fib.HasValidClx)
            return ExtractTextSimple(main);

        return ExtractTextFromClx(main, fib.SelectedTableStream, fib.FcClx, fib.LcbClx);
    }

    /// <summary>
    /// Fallback for documents without CLX: text starts at byte 0 of main stream.
    /// Uses ccpText as the character count.
    /// </summary>
    private static HWPFTextModel ExtractTextSimple(byte[] main)
    {
        var ccpText = ReadInt32(main, FibOffsetCcpText);
        if (ccpText <= 0) return HWPFTextModel.Empty;
        // Word 97+ non-complex documents store text as UTF-16LE starting at offset 0
        var byteCount = Math.Min(ccpText * 2, main.Length);
        try
        {
            var text = Encoding.Unicode.GetString(main, 0, byteCount & ~1);
            return new HWPFTextModel(text, [new HWPFTextPiece(0, text.Length, 0, false)]);
        }
        catch { return HWPFTextModel.Empty; }
    }

    /// <summary>
    /// Parses the CLX in the table stream and assembles text from the piece table.
    /// Ported from ComplexFileTable + TextPieceTable.
    /// </summary>
    private static HWPFTextModel ExtractTextFromClx(byte[] main, byte[] table, int fcClx, int lcbClx)
    {
        var clx = table.AsSpan(fcClx, lcbClx);
        int pos = 0;

        // Skip Prc entries (clxt = 1) — property change records
        while (pos < clx.Length && clx[pos] == 1)
        {
            pos++;                                     // skip clxt byte
            if (pos + 2 > clx.Length) return HWPFTextModel.Empty;
            var cbGrpprl = BinaryPrimitives.ReadUInt16LittleEndian(clx.Slice(pos));
            pos += 2 + cbGrpprl;                       // skip cbGrpprl + grpPrl
        }

        // Expect Pcdt (clxt = 2)
        if (pos >= clx.Length || clx[pos] != 2) return HWPFTextModel.Empty;
        pos++;                                         // skip clxt byte

        if (pos + 4 > clx.Length) return HWPFTextModel.Empty;
        var lcbPcd = BinaryPrimitives.ReadInt32LittleEndian(clx.Slice(pos));
        pos += 4;

        if (lcbPcd <= 0 || pos + lcbPcd > clx.Length) return HWPFTextModel.Empty;
        var plcPcd = clx.Slice(pos, lcbPcd);

        // PlcPcd: (n+1) × int32 CP values, then n × 8-byte PCD structs
        // 4*(n+1) + 8*n = lcbPcd  →  n = (lcbPcd - 4) / 12
        var pieceCount = (lcbPcd - 4) / 12;
        if (pieceCount <= 0) return HWPFTextModel.Empty;

        var sb = new StringBuilder();
        var pieces = new List<HWPFTextPiece>(pieceCount);

        for (int i = 0; i < pieceCount; i++)
        {
            var cpStart = BinaryPrimitives.ReadInt32LittleEndian(plcPcd.Slice(i * 4));
            var cpEnd   = BinaryPrimitives.ReadInt32LittleEndian(plcPcd.Slice((i + 1) * 4));
            var charCount = cpEnd - cpStart;
            if (charCount <= 0) continue;

            // PCD layout within PlcPcd: starts after (pieceCount+1) CPs.
            // PCD struct (8 bytes): descriptor(2 bytes) | fc(4 bytes) | prm(2 bytes)
            // fc is at byte offset +2 within the PCD — ported from PieceDescriptor.java:
            //   descriptor = LittleEndian.getShort(buf, offset);
            //   fc = LittleEndian.getInt(buf, offset + 2);
            var pcdOffset = (pieceCount + 1) * 4 + i * 8;
            if (pcdOffset + 8 > plcPcd.Length) break;

            var fc = BinaryPrimitives.ReadInt32LittleEndian(plcPcd.Slice(pcdOffset + 2));
            bool compressed = (fc & 0x40000000) != 0;
            int filePos;
            if (compressed)
            {
                // CP1252: POI does fc &= ~0x40000000; fc /= 2; → byte offset = fc_val/2
                filePos = (fc & 0x3FFFFFFF) / 2;
            }
            else
            {
                // Unicode: POI returns fc directly without division.
                filePos = fc & 0x3FFFFFFF;
            }

            var textStart = sb.Length;
            AppendPieceText(sb, main, filePos, charCount, compressed);
            var textEnd = sb.Length;
            if (textEnd > textStart)
                pieces.Add(new HWPFTextPiece(textStart, textEnd, filePos, compressed));
        }

        return new HWPFTextModel(sb.ToString(), pieces);
    }

    private static void AppendPieceText(StringBuilder sb, byte[] main, int filePos, int charCount, bool compressed)
    {
        try
        {
            if (compressed)
            {
                var byteCount = Math.Min(charCount, main.Length - filePos);
                if (byteCount <= 0) return;
                var text = LocaleUtil1252.GetString(main, filePos, byteCount);
                sb.Append(text);
            }
            else
            {
                var byteCount = Math.Min(charCount * 2, main.Length - filePos);
                if (byteCount <= 0) return;
                sb.Append(Encoding.Unicode.GetString(main, filePos, byteCount & ~1));
            }
        }
        catch { /* skip corrupt pieces */ }
    }

    // ─── CHPX / formatting ──────────────────────────────────────────────────────

    private static string[] ReadFontTable(byte[] main, IReadOnlyDictionary<string, byte[]> streams, HWPFFileInformationBlock fib)
    {
        var fc = ReadInt32Safe(main, FibOffsetFcSttbfffn);
        var lcb = ReadInt32Safe(main, FibOffsetLcbSttbfffn);
        if (lcb <= 4 || fib.SelectedTableStream is null || fc < 0 || fc + lcb > fib.SelectedTableStream.Length)
            return Array.Empty<string>();

        var tbl = fib.SelectedTableStream.AsSpan(fc, lcb);
        // STTBFFFN: short cstd | short cbExtra | FFN[0..cstd-1]
        // FFN fixed header: cbFfnM1(1) + info(1) + wWeight(2) + chs(1) + ixchSzAlt(1) + panose(10) + fontSig(24) = 40 bytes
        // Font name (Unicode, null-terminated) starts at offset 40 within each FFN
        if (tbl.Length < 4) return Array.Empty<string>();
        var cstd = BinaryPrimitives.ReadInt16LittleEndian(tbl);
        if (cstd < 0 || cstd > 256) cstd = 16;
        var pos = 4; // skip cstd + cbExtra
        var names = new List<string>(cstd);
        for (var i = 0; i < cstd && pos < tbl.Length; i++)
        {
            var cbFfnM1 = tbl[pos];
            var ffnLen = cbFfnM1 + 1;
            if (pos + ffnLen > tbl.Length) break;
            const int nameOffset = 40; // fixed header size per POI Ffn.java
            if (nameOffset < ffnLen)
            {
                var nameStart = pos + nameOffset;
                var nameEnd = nameStart;
                while (nameEnd + 1 < pos + ffnLen && !(tbl[nameEnd] == 0 && tbl[nameEnd + 1] == 0))
                    nameEnd += 2;
                names.Add(nameEnd > nameStart
                    ? Encoding.Unicode.GetString(tbl.Slice(nameStart, nameEnd - nameStart).ToArray())
                    : string.Empty);
            }
            else names.Add(string.Empty);
            pos += ffnLen;
        }
        return names.ToArray();
    }

    /// <summary>
    /// Reads all CHPFKP runs and returns a sorted list of (CP_start, CP_end, CHPX) segments.
    /// Ported from org.apache.poi.hwpf.model.CHPBinTable / CHPFKP.
    /// </summary>
    private static List<(int CpStart, int CpEnd, HWPFChpProperties Chp)> ReadChpSegments(
        byte[] main, HWPFFileInformationBlock fib, HWPFTextModel model, string[] fontTable)
    {
        var result = new List<(int, int, HWPFChpProperties)>();

        var fcChpx = ReadInt32Safe(main, FibOffsetFcPlcfBteChpx);
        var lcbChpx = ReadInt32Safe(main, FibOffsetLcbPlcfBteChpx);
        if (lcbChpx <= 4 || fib.SelectedTableStream is null || fcChpx < 0 || fcChpx + lcbChpx > fib.SelectedTableStream.Length)
            return result;

        var tbl = fib.SelectedTableStream.AsSpan(fcChpx, lcbChpx);
        var n = (lcbChpx - 4) / 8;
        if (n <= 0) return result;

        var fcStarts = new int[n + 1];
        for (var i = 0; i <= n; i++)
            fcStarts[i] = BinaryPrimitives.ReadInt32LittleEndian(tbl.Slice(i * 4));
        var pages = new uint[n];
        for (var i = 0; i < n; i++)
            pages[i] = BinaryPrimitives.ReadUInt32LittleEndian(tbl.Slice((n + 1) * 4 + i * 4));

        for (var pageIdx = 0; pageIdx < n; pageIdx++)
        {
            var pageOffset = (int)pages[pageIdx] * 512;
            if (pageOffset < 0 || pageOffset + 512 > main.Length) continue;

            var fkp = main.AsSpan(pageOffset, 512);
            var crun = fkp[511];
            if (crun == 0 || 4 * (crun + 1) > 512) continue;

            var rgbOffset = 4 * (crun + 1);
            for (var r = 0; r < crun; r++)
            {
                var runFcStart = BinaryPrimitives.ReadInt32LittleEndian(fkp.Slice(r * 4));
                var runFcEnd   = BinaryPrimitives.ReadInt32LittleEndian(fkp.Slice((r + 1) * 4));

                // Convert FC → CP using piece table
                var cpStart = FcToCp(runFcStart, model);
                var cpEnd   = FcToCp(runFcEnd,   model);
                if (cpStart < 0 || cpEnd <= cpStart) continue;

                HWPFChpProperties chp;
                var chpxByteOff = fkp[rgbOffset + r];
                if (chpxByteOff == 0)
                {
                    chp = HWPFChpProperties.Default;
                }
                else
                {
                    var chpxAbsolute = chpxByteOff * 2;
                    if (chpxAbsolute >= 512) { chp = HWPFChpProperties.Default; }
                    else
                    {
                        var cbChpx = fkp[chpxAbsolute];
                        if (chpxAbsolute + 1 + cbChpx > 512) { chp = HWPFChpProperties.Default; }
                        else chp = HWPFChpProperties.ParseChpx(fkp.Slice(chpxAbsolute + 1, cbChpx), fontTable);
                    }
                }
                result.Add((cpStart, cpEnd, chp));
            }
        }
        result.Sort((a, b) => a.Item1.CompareTo(b.Item1));
        return result;
    }

    private static int FcToCp(int fc, HWPFTextModel model)
    {
        foreach (var piece in model.Pieces)
        {
            int pieceByteLen = piece.Compressed ? (piece.End - piece.Start) : (piece.End - piece.Start) * 2;
            int pieceFileEnd = piece.FileOffset + pieceByteLen;
            if (fc < piece.FileOffset || fc > pieceFileEnd) continue;
            int offsetFromStart = fc - piece.FileOffset;
            int cpOffset = piece.Compressed ? offsetFromStart : offsetFromStart / 2;
            return piece.Start + cpOffset;
        }
        return -1;
    }

    // Keep for backward compat — now unused internally
    private static HWPFChpProperties[] ReadPieceChpProps(
        byte[] main, HWPFFileInformationBlock fib, HWPFTextModel model, string[] fontTable)
        => Array.Empty<HWPFChpProperties>();

    private static int ReadInt32(byte[] data, int offset) =>
        BinaryPrimitives.ReadInt32LittleEndian(data.AsSpan(offset));

    private static void WriteInt32(byte[] data, int offset, int value) =>
        BinaryPrimitives.WriteInt32LittleEndian(data.AsSpan(offset), value);

    internal static int ReadInt32Safe(byte[] data, int offset) =>
        offset >= 0 && offset + sizeof(int) <= data.Length ? ReadInt32(data, offset) : 0;

    internal static ushort ReadUInt16Safe(byte[] data, int offset) =>
        offset >= 0 && offset + sizeof(ushort) <= data.Length
            ? BinaryPrimitives.ReadUInt16LittleEndian(data.AsSpan(offset))
            : (ushort)0;

    internal static string GetDeclaredTableStreamName(byte[] main)
    {
        var flags1 = ReadUInt16Safe(main, FibOffsetFlags1);
        return ((flags1 >> 9) & 1) == 1 ? Stream1Table : Stream0Table;
    }

    internal static string GetFallbackTableStreamName(string tableStreamName) =>
        tableStreamName == Stream1Table ? Stream0Table : Stream1Table;

    internal static int GetFibCcpText(byte[] main) => ReadInt32Safe(main, FibOffsetCcpText);

    internal static int GetFibFcClx(byte[] main) => ReadInt32Safe(main, FibOffsetFcClx);

    internal static int GetFibLcbClx(byte[] main) => ReadInt32Safe(main, FibOffsetLcbClx);

    private void SetMainBodyText(string text)
    {
        if (_fib.SelectedTableStreamName is null)
            throw new NotSupportedException("Cannot edit HWPF document without a selected table stream.");

        var main = _mainStream.ToArray();
        var textOffset = AlignEven(main.Length);
        var textBytes = Encoding.Unicode.GetBytes(text);
        Array.Resize(ref main, textOffset + textBytes.Length);
        Buffer.BlockCopy(textBytes, 0, main, textOffset, textBytes.Length);

        var table = _fileSystem.Streams[_fib.SelectedTableStreamName].ToArray();
        var clxOffset = table.Length;
        var clx = BuildSingleUnicodePieceClx(text.Length, textOffset);
        Array.Resize(ref table, clxOffset + clx.Length);
        Buffer.BlockCopy(clx, 0, table, clxOffset, clx.Length);

        WriteInt32(main, FibOffsetCcpText, text.Length);
        WriteInt32(main, FibOffsetFcClx, clxOffset);
        WriteInt32(main, FibOffsetLcbClx, clx.Length);

        _fileSystem.Streams[StreamWordDocument] = main;
        _fileSystem.Streams[_fib.SelectedTableStreamName] = table;
        _mainStream = main;
        _fib = HWPFFileInformationBlock.Read(_mainStream, _fileSystem.Streams);
        _textModel = ExtractTextModel(_mainStream, _fib);
        _text = _textModel.Text;
    }

    private static byte[] BuildSingleUnicodePieceClx(int charCount, int fileOffset)
    {
        const int pieceCount = 1;
        const int plcPcdLength = 4 * (pieceCount + 1) + 8 * pieceCount;
        var clx = new byte[1 + 4 + plcPcdLength];
        clx[0] = 2; // Pcdt
        BinaryPrimitives.WriteInt32LittleEndian(clx.AsSpan(1), plcPcdLength);
        BinaryPrimitives.WriteInt32LittleEndian(clx.AsSpan(5), 0);
        BinaryPrimitives.WriteInt32LittleEndian(clx.AsSpan(9), charCount);

        var pcdOffset = 5 + 4 * (pieceCount + 1);
        BinaryPrimitives.WriteUInt16LittleEndian(clx.AsSpan(pcdOffset), 0);
        BinaryPrimitives.WriteInt32LittleEndian(clx.AsSpan(pcdOffset + 2), fileOffset);
        BinaryPrimitives.WriteUInt16LittleEndian(clx.AsSpan(pcdOffset + 6), 0);
        return clx;
    }

    private static int AlignEven(int value) => (value + 1) & ~1;

    private static string ReplaceOrdinal(string text, string placeholder, string value)
    {
        var index = text.IndexOf(placeholder, StringComparison.Ordinal);
        if (index < 0) return text;

        var sb = new StringBuilder(text.Length);
        var start = 0;
        while (index >= 0)
        {
            sb.Append(text, start, index - start);
            sb.Append(value);
            start = index + placeholder.Length;
            index = text.IndexOf(placeholder, start, StringComparison.Ordinal);
        }

        sb.Append(text, start, text.Length - start);
        return sb.ToString();
    }
}

/// <summary>
/// Minimal read-only port of org.apache.poi.hwpf.usermodel.Range.
/// It exposes the main text as paragraphs and character runs while HWPF editing
/// and full PAPX/CHPX property expansion are still being ported.
/// </summary>
public class Range
{
    private readonly HWPFTextModel _model;
    internal readonly List<(int CpStart, int CpEnd, HWPFChpProperties Chp)> _chpSegments;
    private IReadOnlyList<Paragraph>? _paragraphs;
    private IReadOnlyList<CharacterRun>? _characterRuns;

    internal Range(int start, int end, HWPFTextModel model,
        List<(int CpStart, int CpEnd, HWPFChpProperties Chp)>? chpSegments = null)
    {
        _model = model;
        _chpSegments = chpSegments ?? new List<(int, int, HWPFChpProperties)>();
        Start = Clamp(start, 0, model.Text.Length);
        End = Clamp(end, Start, model.Text.Length);
    }

    internal int Start { get; }

    internal int End { get; }

    public int getStartOffset() => Start;

    public int getEndOffset() => End;

    public string text() => _model.Text.Substring(Start, End - Start);

    public int numParagraphs() => GetParagraphs().Count;

    public Paragraph getParagraph(int index) => GetParagraphs()[index];

    public int numCharacterRuns() => GetCharacterRuns().Count;

    public CharacterRun getCharacterRun(int index) => GetCharacterRuns()[index];

    internal IReadOnlyList<Paragraph> GetParagraphs()
    {
        if (_paragraphs is not null) return _paragraphs;

        var paragraphs = new List<Paragraph>();
        var text = _model.Text;
        var paragraphStart = Start;
        for (var i = Start; i < End; i++)
        {
            if (!IsParagraphTerminator(text[i])) continue;
            paragraphs.Add(new Paragraph(paragraphStart, i + 1, _model, _chpSegments));
            paragraphStart = i + 1;
        }

        if (paragraphStart < End || paragraphs.Count == 0 && Start == End)
            paragraphs.Add(new Paragraph(paragraphStart, End, _model, _chpSegments));

        _paragraphs = paragraphs;
        return _paragraphs;
    }

    internal IReadOnlyList<CharacterRun> GetCharacterRuns()
    {
        if (_characterRuns is not null) return _characterRuns;

        var runs = new List<CharacterRun>();
        if (_chpSegments.Count > 0)
        {
            // Build runs from CHPFKP segment boundaries intersected with this Range,
            // filling gaps with default-properties runs so that all text is covered.
            var cursor = Start;
            foreach (var seg in _chpSegments)
            {
                var segStart = Math.Max(Start, seg.CpStart);
                var segEnd   = Math.Min(End, seg.CpEnd);
                if (segStart >= End) break;
                if (segEnd <= Start) continue;
                // Fill gap before this segment
                if (cursor < segStart)
                    runs.Add(new CharacterRun(cursor, segStart, _model, FindPiece(cursor), HWPFChpProperties.Default));
                // Add this segment
                runs.Add(new CharacterRun(segStart, segEnd, _model, FindPiece(segStart), seg.Chp));
                cursor = Math.Max(cursor, segEnd);
            }
            // Fill gap after last segment
            if (cursor < End)
                runs.Add(new CharacterRun(cursor, End, _model, FindPiece(cursor), HWPFChpProperties.Default));
        }
        else
        {
            // Fallback: one run per text piece
            foreach (var piece in _model.Pieces)
            {
                var start = Math.Max(Start, piece.Start);
                var end = Math.Min(End, piece.End);
                if (start < end)
                    runs.Add(new CharacterRun(start, end, _model, piece, HWPFChpProperties.Default));
            }
        }

        if (runs.Count == 0 && Start < End)
            runs.Add(new CharacterRun(Start, End, _model, null, HWPFChpProperties.Default));

        _characterRuns = runs;
        return _characterRuns;
    }

    private HWPFTextPiece? FindPiece(int cpStart)
    {
        foreach (var piece in _model.Pieces)
            if (cpStart >= piece.Start && cpStart < piece.End) return piece;
        return null;
    }

    private static bool IsParagraphTerminator(char ch) =>
        ch is '\r' or '\f' or '\a';

    private static int Clamp(int value, int min, int max)
    {
        if (value < min) return min;
        return value > max ? max : value;
    }
}

/// <summary>Minimal read-only port of org.apache.poi.hwpf.usermodel.Paragraph.</summary>
public sealed class Paragraph : Range
{
    internal Paragraph(int start, int end, HWPFTextModel model,
        List<(int CpStart, int CpEnd, HWPFChpProperties Chp)>? chpSegments = null)
        : base(start, end, model, chpSegments)
    {
    }

    public int getIndentFromLeft() => 0;

    public int getIndentFromRight() => 0;

    public int getFirstLineIndent() => 0;

    public int getJustification() => 0;
}

/// <summary>Minimal read-only port of org.apache.poi.hwpf.usermodel.CharacterRun.</summary>
public sealed class CharacterRun : Range
{
    private readonly HWPFTextPiece? _piece;
    private readonly HWPFChpProperties _chp;

    internal CharacterRun(int start, int end, HWPFTextModel model, HWPFTextPiece? piece, HWPFChpProperties chp)
        : base(start, end, model, null)
    {
        _piece = piece;
        _chp = chp;
    }

    public bool isBold() => _chp.Bold;

    public bool isItalic() => _chp.Italic;

    public bool isStrikeThrough() => _chp.Strike;

    public int getUnderlineCode() => _chp.Underline;

    /// <summary>Font size in half-points (e.g., 22 = 11pt). Returns 0 if unknown.</summary>
    public int getFontSize() => _chp.FontSizeHalfPoints;

    public string getFontName() => _chp.FontName;

    public bool isCompressed() => _piece?.Compressed ?? false;
}

internal sealed class HWPFTextModel
{
    internal static HWPFTextModel Empty { get; } = new(string.Empty, []);

    internal HWPFTextModel(string text, IReadOnlyList<HWPFTextPiece> pieces)
    {
        Text = text;
        Pieces = pieces;
    }

    internal string Text { get; }

    internal IReadOnlyList<HWPFTextPiece> Pieces { get; }
}

internal sealed record HWPFTextPiece(int Start, int End, int FileOffset, bool Compressed);

public sealed class HWPFFileInformationBlock
{
    private HWPFFileInformationBlock(
        bool fWhichTblStm,
        int ccpText,
        int fcClx,
        int lcbClx,
        string declaredTableStreamName,
        string? selectedTableStreamName,
        bool usedTableStreamFallback,
        byte[]? selectedTableStream)
    {
        FWhichTblStm = fWhichTblStm;
        CcpText = ccpText;
        FcClx = fcClx;
        LcbClx = lcbClx;
        DeclaredTableStreamName = declaredTableStreamName;
        SelectedTableStreamName = selectedTableStreamName;
        UsedTableStreamFallback = usedTableStreamFallback;
        SelectedTableStream = selectedTableStream;
        SelectedTableStreamLength = selectedTableStream?.Length ?? 0;
    }

    public bool FWhichTblStm { get; }

    public int CcpText { get; }

    public int FcClx { get; }

    public int LcbClx { get; }

    public string DeclaredTableStreamName { get; }

    public string? SelectedTableStreamName { get; }

    public bool UsedTableStreamFallback { get; }

    public int SelectedTableStreamLength { get; }

    public bool HasSelectedTableStream => SelectedTableStream is not null;

    public bool HasValidClx =>
        SelectedTableStream is not null &&
        FcClx >= 0 &&
        LcbClx > 0 &&
        FcClx <= SelectedTableStream.Length &&
        LcbClx <= SelectedTableStream.Length - FcClx;

    internal byte[]? SelectedTableStream { get; }

    internal static HWPFFileInformationBlock Read(byte[] main, IReadOnlyDictionary<string, byte[]> streams)
    {
        var declaredTableStreamName = HWPFDocument.GetDeclaredTableStreamName(main);
        var selectedTableStreamName = declaredTableStreamName;
        var usedFallback = false;
        if (!streams.TryGetValue(selectedTableStreamName, out var tableStream))
        {
            selectedTableStreamName = HWPFDocument.GetFallbackTableStreamName(declaredTableStreamName);
            usedFallback = streams.TryGetValue(selectedTableStreamName, out tableStream);
            if (!usedFallback)
            {
                selectedTableStreamName = null;
            }
        }

        return new HWPFFileInformationBlock(
            declaredTableStreamName == "1Table",
            HWPFDocument.GetFibCcpText(main),
            HWPFDocument.GetFibFcClx(main),
            HWPFDocument.GetFibLcbClx(main),
            declaredTableStreamName,
            selectedTableStreamName,
            usedFallback,
            tableStream);
    }
}

/// <summary>Character properties parsed from CHPX sprms.</summary>
internal sealed class HWPFChpProperties
{
    internal static HWPFChpProperties Default { get; } = new();

    internal bool Bold { get; set; }
    internal bool Italic { get; set; }
    internal bool Strike { get; set; }
    internal byte Underline { get; set; }
    internal int FontSizeHalfPoints { get; set; } // 0 = unknown
    internal string FontName { get; set; } = string.Empty;

    internal static HWPFChpProperties ParseChpx(ReadOnlySpan<byte> chpx, string[] fontTable)
    {
        var props = new HWPFChpProperties();
        var pos = 0;
        while (pos + 2 <= chpx.Length)
        {
            var sprm = BinaryPrimitives.ReadUInt16LittleEndian(chpx.Slice(pos));
            pos += 2;
            // spra = bits 13-15 of sprm
            var spra = (sprm >> 13) & 0x07;
            int opSize;
            switch (spra)
            {
                case 0: case 1: opSize = 1; break;
                case 2: case 4: opSize = 2; break;
                case 3: opSize = 4; break;
                case 7: opSize = 3; break;
                case 5: // variable: cbGrpprl byte + cbGrpprl bytes
                    opSize = pos < chpx.Length ? chpx[pos] + 1 : 0;
                    break;
                case 6: // variable complex
                    opSize = pos < chpx.Length ? chpx[pos] + 1 : 0;
                    break;
                default: return props; // unknown
            }
            if (pos + opSize > chpx.Length) break;
            var operand = chpx.Slice(pos, opSize);
            pos += opSize;

            switch (sprm)
            {
                case 0x0835: props.Bold = operand[0] != 0 && operand[0] != 128; break; // sprmCFBold
                case 0x0836: props.Italic = operand[0] != 0 && operand[0] != 128; break; // sprmCFItalic
                case 0x0837: props.Strike = operand[0] != 0 && operand[0] != 128; break; // sprmCFStrike
                case 0x2A3E: props.Underline = operand[0]; break; // sprmCKul
                case 0x4A43: // sprmCHps: font size in half-points
                    props.FontSizeHalfPoints = BinaryPrimitives.ReadUInt16LittleEndian(operand);
                    break;
                case 0x4A4F: case 0x4A51: case 0x4A53: // sprmCRgFtc0/1/2: font index
                    if (sprm == 0x4A4F) // ASCII font
                    {
                        var idx = BinaryPrimitives.ReadUInt16LittleEndian(operand);
                        if (idx < fontTable.Length) props.FontName = fontTable[idx];
                    }
                    break;
            }
        }
        return props;
    }
}

/// <summary>Lazy-initialized Windows-1252 encoding.</summary>
internal static class LocaleUtil1252
{
    private static Encoding? _enc;

    internal static string GetString(byte[] data, int offset, int count)
    {
        if (_enc is null)
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            _enc = Encoding.GetEncoding(1252);
        }
        return _enc.GetString(data, offset, count);
    }
}
