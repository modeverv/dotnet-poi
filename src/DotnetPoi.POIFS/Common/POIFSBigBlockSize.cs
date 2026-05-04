using DotnetPoi.SS.Util;

namespace DotnetPoi.POIFS.Common;

public sealed class POIFSBigBlockSize
{
    private readonly int _bigBlockSize;
    private readonly short _headerValue;

    internal POIFSBigBlockSize(int bigBlockSize, short headerValue)
    {
        _bigBlockSize = bigBlockSize;
        _headerValue = headerValue;
    }

    public int GetBigBlockSize()
    {
        return _bigBlockSize;
    }

    public short GetHeaderValue()
    {
        return _headerValue;
    }

    public int GetPropertiesPerBlock()
    {
        return _bigBlockSize / POIFSConstants.PropertySize;
    }

    public int GetBatEntriesPerBlock()
    {
        return _bigBlockSize / LittleEndianConsts.IntSize;
    }

    public int GetXbatEntriesPerBlock()
    {
        return GetBatEntriesPerBlock() - 1;
    }

    public int GetNextXbatChainOffset()
    {
        return GetXbatEntriesPerBlock() * LittleEndianConsts.IntSize;
    }
}

