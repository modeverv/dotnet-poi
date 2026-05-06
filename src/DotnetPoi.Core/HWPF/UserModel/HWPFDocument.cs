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
    private const int FibOffsetFlags1     = 10;   // 2-byte flags: bit9=fWhichTblStm
    private const int FibOffsetCcpText    = 72;   // fibRgLw97 base (60) + 0xC = 72
    private const int FibOffsetFcClx      = 418;  // fibRgFcLcb start (154) + CLX index 33 * 8
    private const int FibOffsetLcbClx     = 422;  // fcClx + 4

    private readonly byte[] _mainStream;
    private readonly string _text;

    public HWPFDocument(Stream stream)
    {
        Guard.ThrowIfNull(stream, nameof(stream));
        var streams = CompoundFile.ReadStreams(stream);
        if (!streams.TryGetValue(StreamWordDocument, out var mainStream))
            throw new InvalidDataException("Not a valid .doc file: missing WordDocument stream.");
        _mainStream = mainStream;
        _text = ExtractText(mainStream, streams);
    }

    /// <summary>
    /// Returns the plain text of the main document body.
    /// Ported from HWPFDocument.getText().
    /// </summary>
    public string getText() => _text;

    /// <summary>Returns the character count of the main document body.</summary>
    public int getCcpText() => ReadInt32(_mainStream, FibOffsetCcpText);

    public void Dispose() { }

    // ─── text extraction ────────────────────────────────────────────────────

    private static string ExtractText(byte[] main, Dictionary<string, byte[]> streams)
    {
        // Determine which table stream to use (fWhichTblStm bit 9 of flags1)
        var flags1 = BinaryPrimitives.ReadUInt16LittleEndian(main.AsSpan(FibOffsetFlags1));
        var whichTable = (flags1 >> 9) & 1;
        var tableStreamName = whichTable == 1 ? Stream1Table : Stream0Table;

        if (!streams.TryGetValue(tableStreamName, out var tableStream))
        {
            // Fall back to the other table
            tableStreamName = whichTable == 1 ? Stream0Table : Stream1Table;
            if (!streams.TryGetValue(tableStreamName, out tableStream))
                return string.Empty;
        }

        var fcClx  = ReadInt32(main, FibOffsetFcClx);
        var lcbClx = ReadInt32(main, FibOffsetLcbClx);
        if (fcClx < 0 || lcbClx <= 0 || fcClx + lcbClx > tableStream.Length)
            return ExtractTextSimple(main);

        return ExtractTextFromClx(main, tableStream, fcClx, lcbClx);
    }

    /// <summary>
    /// Fallback for documents without CLX: text starts at byte 0 of main stream.
    /// Uses ccpText as the character count.
    /// </summary>
    private static string ExtractTextSimple(byte[] main)
    {
        var ccpText = ReadInt32(main, FibOffsetCcpText);
        if (ccpText <= 0) return string.Empty;
        // Word 97+ non-complex documents store text as UTF-16LE starting at offset 0
        var byteCount = Math.Min(ccpText * 2, main.Length);
        try { return Encoding.Unicode.GetString(main, 0, byteCount); }
        catch { return string.Empty; }
    }

    /// <summary>
    /// Parses the CLX in the table stream and assembles text from the piece table.
    /// Ported from ComplexFileTable + TextPieceTable.
    /// </summary>
    private static string ExtractTextFromClx(byte[] main, byte[] table, int fcClx, int lcbClx)
    {
        var clx = table.AsSpan(fcClx, lcbClx);
        int pos = 0;

        // Skip Prc entries (clxt = 1) — property change records
        while (pos < clx.Length && clx[pos] == 1)
        {
            pos++;                                     // skip clxt byte
            if (pos + 2 > clx.Length) return string.Empty;
            var cbGrpprl = BinaryPrimitives.ReadUInt16LittleEndian(clx.Slice(pos));
            pos += 2 + cbGrpprl;                       // skip cbGrpprl + grpPrl
        }

        // Expect Pcdt (clxt = 2)
        if (pos >= clx.Length || clx[pos] != 2) return string.Empty;
        pos++;                                         // skip clxt byte

        if (pos + 4 > clx.Length) return string.Empty;
        var lcbPcd = BinaryPrimitives.ReadInt32LittleEndian(clx.Slice(pos));
        pos += 4;

        if (lcbPcd <= 0 || pos + lcbPcd > clx.Length) return string.Empty;
        var plcPcd = clx.Slice(pos, lcbPcd);

        // PlcPcd: (n+1) × int32 CP values, then n × 8-byte PCD structs
        // 4*(n+1) + 8*n = lcbPcd  →  n = (lcbPcd - 4) / 12
        var pieceCount = (lcbPcd - 4) / 12;
        if (pieceCount <= 0) return string.Empty;

        var sb = new StringBuilder();

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

            AppendPieceText(sb, main, filePos, charCount, compressed);
        }

        return sb.ToString();
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

    private static int ReadInt32(byte[] data, int offset) =>
        BinaryPrimitives.ReadInt32LittleEndian(data.AsSpan(offset));
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
