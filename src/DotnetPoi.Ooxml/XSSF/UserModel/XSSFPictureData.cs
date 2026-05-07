using DotnetPoi.SS.UserModel;

namespace DotnetPoi.XSSF.UserModel;

public sealed class XSSFPictureData : IPictureData
{
    internal XSSFPictureData(byte[] data, int pictureType, int index)
    {
        Data = data.ToArray();
        PictureType = pictureType;
        Index = index;
    }

    internal byte[] Data { get; }

    internal int Index { get; }

    internal int PictureType { get; }

    internal string Extension => PictureType switch
    {
        XSSFWorkbook.PICTURE_TYPE_JPEG => "jpeg",
        XSSFWorkbook.PICTURE_TYPE_PNG => "png",
        XSSFWorkbook.PICTURE_TYPE_DIB => "dib",
        XSSFWorkbook.PICTURE_TYPE_GIF => "gif",
        XSSFWorkbook.PICTURE_TYPE_TIFF => "tiff",
        XSSFWorkbook.PICTURE_TYPE_EPS => "eps",
        XSSFWorkbook.PICTURE_TYPE_BMP => "bmp",
        XSSFWorkbook.PICTURE_TYPE_WPG => "wpg",
        XSSFWorkbook.PICTURE_TYPE_EMF => "emf",
        _ => throw new NotSupportedException($"Picture type {PictureType} is not supported by the Phase 2.5 XSSF writer.")
    };

    internal string ContentType => PictureType switch
    {
        XSSFWorkbook.PICTURE_TYPE_JPEG => "image/jpeg",
        XSSFWorkbook.PICTURE_TYPE_PNG => "image/png",
        XSSFWorkbook.PICTURE_TYPE_DIB => "image/dib",
        XSSFWorkbook.PICTURE_TYPE_GIF => "image/gif",
        XSSFWorkbook.PICTURE_TYPE_TIFF => "image/tiff",
        XSSFWorkbook.PICTURE_TYPE_EPS => "image/eps",
        XSSFWorkbook.PICTURE_TYPE_BMP => "image/bmp",
        XSSFWorkbook.PICTURE_TYPE_WPG => "image/x-wpg",
        XSSFWorkbook.PICTURE_TYPE_EMF => "image/x-emf",
        _ => throw new NotSupportedException($"Picture type {PictureType} is not supported by the Phase 2.5 XSSF writer.")
    };

    public byte[] getData()
    {
        return Data.ToArray();
    }

    public string suggestFileExtension()
    {
        return Extension;
    }

    public int getPictureType()
    {
        return PictureType;
    }

    public string getMimeType()
    {
        return ContentType;
    }
}
