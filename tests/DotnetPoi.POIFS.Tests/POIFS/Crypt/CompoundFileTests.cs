using DotnetPoi.POIFS.Crypt;
using Xunit;

namespace DotnetPoi.POIFS.Tests.Crypt;

public sealed class CompoundFileTests
{
    [Fact]
    public void WriteRead_LargeStream_UsesDifatExtensionSectors()
    {
        var large = new byte[8 * 1024 * 1024 + 123];
        for (var i = 0; i < large.Length; i++)
        {
            large[i] = (byte)(i * 31);
        }

        using var output = new MemoryStream();
        CompoundFile.Write(output, new Dictionary<string, byte[]>(StringComparer.Ordinal)
        {
            ["LargeStream"] = large
        });

        var bytes = output.ToArray();
        Assert.True(BitConverter.ToInt32(bytes, 72) > 0);
        Assert.NotEqual(unchecked((int)0xFFFFFFFE), BitConverter.ToInt32(bytes, 68));

        output.Position = 0;
        var streams = CompoundFile.ReadStreams(output);

        Assert.True(streams.ContainsKey("LargeStream"));
        Assert.Equal(large, streams["LargeStream"]);
    }
}
