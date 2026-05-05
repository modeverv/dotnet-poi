namespace DotnetPoi.POIFS.Common;

public static class POIFSConstants
{
    public const int SmallerBigBlockSize = 0x0200;
    public static readonly POIFSBigBlockSize SmallerBigBlockSizeDetails =
        new(SmallerBigBlockSize, 9);

    public const int LargerBigBlockSize = 0x1000;
    public static readonly POIFSBigBlockSize LargerBigBlockSizeDetails =
        new(LargerBigBlockSize, 12);

    public const int SmallBlockSize = 0x0040;
    public const int PropertySize = 0x0080;
    public const int BigBlockMinimumDocumentSize = 0x1000;

    public const int LargestRegularSectorNumber = -5;
    public const int DifatSectorBlock = -4;
    public const int FatSectorBlock = -3;
    public const int EndOfChain = -2;
    public const int UnusedBlock = -1;
}

