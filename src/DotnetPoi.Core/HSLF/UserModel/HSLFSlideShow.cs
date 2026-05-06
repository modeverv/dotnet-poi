using System.Buffers.Binary;
using System.Text;
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
    private const ushort RecTypeSlide          = 1006;
    private const ushort RecTypeTextCharsAtom  = 4000;  // UTF-16LE text
    private const ushort RecTypeTextBytesAtom  = 4008;  // CP1252 text
    private const byte   ContainerVersionFlag  = 0x0F;  // low nibble of verAndInstance

    private readonly List<HSLFSlide> _slides = new();

    public HSLFSlideShow(Stream stream)
    {
        Guard.ThrowIfNull(stream, nameof(stream));
        var streams = CompoundFile.ReadStreams(stream);
        if (!streams.TryGetValue(StreamPowerPointDocument, out var pptStream))
            throw new InvalidDataException("Not a valid .ppt file: missing 'PowerPoint Document' stream.");
        ParseRecords(pptStream);
    }

    /// <summary>Returns the slides in the presentation.</summary>
    public IReadOnlyList<HSLFSlide> getSlides() => _slides;

    public void Dispose() { }

    // ─── record parsing ─────────────────────────────────────────────────────

    private void ParseRecords(byte[] data)
    {
        int pos = 0;
        while (pos + 8 <= data.Length)
        {
            var verAndInst = BinaryPrimitives.ReadUInt16LittleEndian(data.AsSpan(pos));
            var recType    = BinaryPrimitives.ReadUInt16LittleEndian(data.AsSpan(pos + 2));
            var recLen     = BinaryPrimitives.ReadInt32LittleEndian(data.AsSpan(pos + 4));
            pos += 8;

            if (recLen < 0 || pos + recLen > data.Length) break;

            var body = data.AsSpan(pos, recLen);
            bool isContainer = (verAndInst & 0x0F) == ContainerVersionFlag;

            if (isContainer && recType == RecTypeSlide)
            {
                var slideTexts = new List<string>();
                ExtractTextsFromRecord(body, slideTexts);
                _slides.Add(new HSLFSlide(slideTexts));
            }

            pos += recLen;
        }
    }

    private static void ExtractTextsFromRecord(ReadOnlySpan<byte> data, List<string> texts)
    {
        int pos = 0;
        while (pos + 8 <= data.Length)
        {
            var verAndInst = BinaryPrimitives.ReadUInt16LittleEndian(data.Slice(pos));
            var recType    = BinaryPrimitives.ReadUInt16LittleEndian(data.Slice(pos + 2));
            var recLen     = BinaryPrimitives.ReadInt32LittleEndian(data.Slice(pos + 4));
            pos += 8;

            if (recLen < 0 || pos + recLen > data.Length) break;

            var body = data.Slice(pos, recLen);

            switch (recType)
            {
                case RecTypeTextCharsAtom:
                    if (body.Length >= 2)
                        texts.Add(Encoding.Unicode.GetString(body.ToArray()));
                    break;

                case RecTypeTextBytesAtom:
                    if (body.Length >= 1)
                        texts.Add(LocaleUtil1252Hslf.GetString(body));
                    break;

                default:
                    if ((verAndInst & 0x0F) == ContainerVersionFlag)
                        ExtractTextsFromRecord(body, texts);
                    break;
            }

            pos += recLen;
        }
    }
}

/// <summary>
/// Represents a single slide in a PowerPoint 97–2003 presentation.
/// Ported from org.apache.poi.hslf.usermodel.HSLFSlide.
/// </summary>
public sealed class HSLFSlide
{
    private readonly IReadOnlyList<string> _texts;

    internal HSLFSlide(IReadOnlyList<string> texts)
    {
        _texts = texts;
    }

    /// <summary>Returns all text runs on this slide.</summary>
    public IReadOnlyList<string> getTextParagraphs() => _texts;

    /// <summary>Returns the title text (first text run), or empty string.</summary>
    public string getTitle() => _texts.Count > 0 ? _texts[0] : string.Empty;
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
