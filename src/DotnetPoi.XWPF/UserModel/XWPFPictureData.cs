namespace DotnetPoi.XWPF.UserModel;

public sealed class XWPFPictureData
{
    public const int PICTURE_TYPE_EMF = 2;
    public const int PICTURE_TYPE_WMF = 3;
    public const int PICTURE_TYPE_PICT = 4;
    public const int PICTURE_TYPE_JPEG = 5;
    public const int PICTURE_TYPE_PNG = 6;
    public const int PICTURE_TYPE_DIB = 7;
    public const int PICTURE_TYPE_GIF = 8;
    public const int PICTURE_TYPE_TIFF = 9;
    public const int PICTURE_TYPE_EPS = 10;
    public const int PICTURE_TYPE_BMP = 11;
    public const int PICTURE_TYPE_WPG = 12;

    internal XWPFPictureData(byte[] data, int format, int index)
    {
        Data = (byte[])data.Clone();
        Format = format;
        Index = index;
    }

    internal byte[] Data { get; }
    internal int Format { get; }
    internal int Index { get; }
    internal string Extension => suggestFileExtension();
    internal string ContentType => getMimeType();

    public byte[] getData() => (byte[])Data.Clone();

    public int getPictureType() => Format;

    public string getFileName() => $"image{Index}.{suggestFileExtension()}";

    public string suggestFileExtension() => Format switch
    {
        PICTURE_TYPE_JPEG => "jpeg",
        PICTURE_TYPE_PNG => "png",
        PICTURE_TYPE_GIF => "gif",
        PICTURE_TYPE_DIB => "dib",
        PICTURE_TYPE_TIFF => "tiff",
        PICTURE_TYPE_EPS => "eps",
        PICTURE_TYPE_BMP => "bmp",
        PICTURE_TYPE_WPG => "wpg",
        PICTURE_TYPE_EMF => "emf",
        PICTURE_TYPE_WMF => "wmf",
        PICTURE_TYPE_PICT => "pict",
        _ => "bin"
    };

    public string getMimeType() => Format switch
    {
        PICTURE_TYPE_JPEG => "image/jpeg",
        PICTURE_TYPE_PNG => "image/png",
        PICTURE_TYPE_GIF => "image/gif",
        PICTURE_TYPE_DIB => "image/dib",
        PICTURE_TYPE_TIFF => "image/tiff",
        PICTURE_TYPE_EPS => "image/x-eps",
        PICTURE_TYPE_BMP => "image/bmp",
        PICTURE_TYPE_WPG => "image/x-wpg",
        PICTURE_TYPE_EMF => "image/x-emf",
        PICTURE_TYPE_WMF => "image/x-wmf",
        PICTURE_TYPE_PICT => "image/pict",
        _ => "application/octet-stream"
    };
}
