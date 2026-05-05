using System.Buffers.Binary;

namespace DotnetPoi.SS.Util;

public static class LittleEndian
{
    public static void PutLong(byte[] data, int offset, long value)
    {
        BinaryPrimitives.WriteInt64LittleEndian(data.AsSpan(offset), value);
    }

    public static void PutInt(byte[] data, int offset, int value)
    {
        BinaryPrimitives.WriteInt32LittleEndian(data.AsSpan(offset), value);
    }

    public static void PutShort(byte[] data, int offset, short value)
    {
        BinaryPrimitives.WriteInt16LittleEndian(data.AsSpan(offset), value);
    }
}

