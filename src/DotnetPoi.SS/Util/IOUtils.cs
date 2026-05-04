using System.Buffers;

namespace DotnetPoi.SS.Util;

public static class IOUtils
{
    public static int ReadFully(Stream stream, byte[] buffer, int offset, int count)
    {
        ArgumentNullException.ThrowIfNull(stream);
        ArgumentNullException.ThrowIfNull(buffer);

        var totalRead = 0;
        while (totalRead < count)
        {
            var read = stream.Read(buffer, offset + totalRead, count - totalRead);
            if (read <= 0)
            {
                break;
            }

            totalRead += read;
        }

        return totalRead == 0 ? -1 : totalRead;
    }

    public static byte[] PeekFirstNBytes(Stream stream, int limit)
    {
        ArgumentNullException.ThrowIfNull(stream);
        if (!stream.CanSeek)
        {
            throw new IOException("PeekFirstNBytes requires a seekable stream.");
        }

        var start = stream.Position;
        var buffer = ArrayPool<byte>.Shared.Rent(limit);
        try
        {
            var read = stream.Read(buffer, 0, limit);
            if (read == 0)
            {
                throw new EndOfStreamException("Stream is empty.");
            }

            if (read < limit)
            {
                Array.Clear(buffer, read, limit - read);
            }

            var result = new byte[limit];
            Buffer.BlockCopy(buffer, 0, result, 0, limit);
            stream.Position = start;
            return result;
        }
        finally
        {
            ArrayPool<byte>.Shared.Return(buffer);
        }
    }
}

