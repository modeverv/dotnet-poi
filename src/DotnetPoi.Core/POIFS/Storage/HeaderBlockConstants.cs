using DotnetPoi.POIFS.Common;
using DotnetPoi.SS.Util;

namespace DotnetPoi.POIFS.Storage;

public static class HeaderBlockConstants
{
    public const long Signature = unchecked((long)0xE11AB1A1E011CFD0UL);
    public const int BatArrayOffset = 0x4c;
    public const int MaxBatsInHeader =
        (POIFSConstants.SmallerBigBlockSize - BatArrayOffset) / LittleEndianConsts.IntSize;

    public const int SignatureOffset = 0;
    public const int BatCountOffset = 0x2C;
    public const int PropertyCountOffset = 0x28;
    public const int PropertyStartOffset = 0x30;
    public const int SbatStartOffset = 0x3C;
    public const int SbatBlockCountOffset = 0x40;
    public const int XbatStartOffset = 0x44;
    public const int XbatCountOffset = 0x48;
}

