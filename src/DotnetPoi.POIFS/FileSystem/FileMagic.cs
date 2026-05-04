using DotnetPoi.POIFS.Storage;
using DotnetPoi.SS.Util;

namespace DotnetPoi.POIFS.FileSystem;

public sealed class FileMagic
{
    public const int MaxPatternLength = 44;

    public static readonly FileMagic Ole2 = new("OLE2", FromLong(HeaderBlockConstants.Signature));
    public static readonly FileMagic Ooxml = new("OOXML", FromInts(0x50, 0x4b, 0x03, 0x04));
    public static readonly FileMagic Xml = new("XML", FromInts(0x3c, 0x3f, 0x78, 0x6d, 0x6c));
    public static readonly FileMagic Biff2 = new("BIFF2", FromInts(0x09, 0x00, 0x04, 0x00, 0x00, 0x00, '?', 0x00));
    public static readonly FileMagic Biff3 = new("BIFF3", FromInts(0x09, 0x02, 0x06, 0x00, 0x00, 0x00, '?', 0x00));
    public static readonly FileMagic Biff4 = new("BIFF4",
        FromBytes(new byte[] { 0x09, 0x04, 0x06, 0x00, 0x00, 0x00, (byte)'?', 0x00 }),
        FromBytes(new byte[] { 0x09, 0x04, 0x06, 0x00, 0x00, 0x00, 0x00, 0x01 }));
    public static readonly FileMagic MsWrite = new("MSWRITE",
        FromBytes(new byte[] { 0x31, 0xbe, 0x00, 0x00 }),
        FromBytes(new byte[] { 0x32, 0xbe, 0x00, 0x00 }));
    public static readonly FileMagic Rtf = new("RTF", FromString("{\\rtf"));
    public static readonly FileMagic Pdf = new("PDF", FromString("%PDF"));
    public static readonly FileMagic Html = new("HTML", FromString("<!DOCTYP"),
        FromString("<html"), FromString("\n\r<html"), FromString("\r\n<html"), FromString("\r<html"), FromString("\n<html"),
        FromString("<HTML"), FromString("\r\n<HTML"), FromString("\n\r<HTML"), FromString("\r<HTML"), FromString("\n<HTML"));
    public static readonly FileMagic Word2 = new("WORD2", FromInts(0xdb, 0xa5, 0x2d, 0x00));
    public static readonly FileMagic Jpeg = new("JPEG",
        FromBytes(new byte[] { 0xFF, 0xD8, 0xFF, 0xDB }),
        FromBytes(new byte[] { 0xFF, 0xD8, 0xFF, 0xE0, (byte)'?', (byte)'?', (byte)'J', (byte)'F', (byte)'I', (byte)'F', 0x00, 0x01 }),
        FromBytes(new byte[] { 0xFF, 0xD8, 0xFF, 0xEE }),
        FromBytes(new byte[] { 0xFF, 0xD8, 0xFF, 0xE1, (byte)'?', (byte)'?', (byte)'E', (byte)'x', (byte)'i', (byte)'f', 0x00, 0x00 }));
    public static readonly FileMagic Gif = new("GIF", FromString("GIF87a"), FromString("GIF89a"));
    public static readonly FileMagic Png = new("PNG", FromInts(0x89, 'P', 'N', 'G', 0x0D, 0x0A, 0x1A, 0x0A));
    public static readonly FileMagic Tiff = new("TIFF", FromString("II*\0"), FromString("MM\0*"));
    public static readonly FileMagic Wmf = new("WMF", FromInts(0xD7, 0xCD, 0xC6, 0x9A));
    public static readonly FileMagic Emf = new("EMF", FromInts(1, 0, 0, 0,
        '?', '?', '?', '?', '?', '?', '?', '?', '?', '?', '?', '?', '?', '?', '?', '?', '?', '?',
        '?', '?', '?', '?', '?', '?', '?', '?', '?', '?', '?', '?', '?', '?', '?', '?', '?', '?',
        ' ', 'E', 'M', 'F'));
    public static readonly FileMagic Bmp = new("BMP", FromInts('B', 'M'));
    public static readonly FileMagic Unknown = new("UNKNOWN", Array.Empty<byte[]>());

    private static readonly FileMagic[] All =
    {
        Ole2, Ooxml, Xml, Biff2, Biff3, Biff4, MsWrite, Rtf, Pdf, Html, Word2, Jpeg, Gif, Png, Tiff, Wmf, Emf, Bmp, Unknown
    };

    private readonly byte[][] _magic;

    private FileMagic(string name, params byte[][] magic)
    {
        Name = name;
        _magic = magic;
    }

    public string Name { get; }

    public IReadOnlyList<byte[]> MagicPatterns => _magic;

    public static IReadOnlyList<FileMagic> Values => All;

    public static FileMagic ValueOf(byte[] magic)
    {
        foreach (var fm in All)
        {
            foreach (var pattern in fm._magic)
            {
                if (magic.Length < pattern.Length)
                {
                    continue;
                }

                if (FindMagic(pattern, magic))
                {
                    return fm;
                }
            }
        }

        return Unknown;
    }

    public static FileMagic ValueOf(FileInfo file)
    {
        using var stream = file.OpenRead();
        var buffer = new byte[MaxPatternLength];
        var read = IOUtils.ReadFully(stream, buffer, 0, buffer.Length);
        if (read == -1)
        {
            return Unknown;
        }

        var data = buffer.AsSpan(0, read).ToArray();
        return ValueOf(data);
    }

    public static FileMagic ValueOf(Stream stream)
    {
        if (!stream.CanSeek)
        {
            throw new IOException("ValueOf requires a seekable stream.");
        }

        var data = IOUtils.PeekFirstNBytes(stream, MaxPatternLength);
        return ValueOf(data);
    }

    public static Stream PrepareToCheckMagic(Stream stream)
    {
        return stream.CanSeek ? stream : new BufferedStream(stream);
    }

    private static bool FindMagic(byte[] expected, byte[] actual)
    {
        for (var i = 0; i < expected.Length; i++)
        {
            var expectedByte = expected[i];
            if (actual[i] != expectedByte && expectedByte != (byte)'?')
            {
                return false;
            }
        }

        return true;
    }

    private static byte[] FromLong(long value)
    {
        var bytes = new byte[8];
        LittleEndian.PutLong(bytes, 0, value);
        return bytes;
    }

    private static byte[] FromInts(params int[] magic)
    {
        var bytes = new byte[magic.Length];
        for (var i = 0; i < magic.Length; i++)
        {
            bytes[i] = (byte)(magic[i] & 0xFF);
        }

        return bytes;
    }

    private static byte[] FromBytes(byte[] bytes)
    {
        return bytes;
    }

    private static byte[] FromString(string value)
    {
        return LocaleUtil.Charset1252.GetBytes(value);
    }
}
