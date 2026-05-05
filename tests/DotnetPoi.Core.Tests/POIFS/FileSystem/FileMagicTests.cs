using System.Text;
using DotnetPoi.POIFS.FileSystem;
using Xunit;

namespace DotnetPoi.POIFS.Tests.FileSystem;

public class FileMagicTests
{
    [Fact]
    public void ValueOf_MagicBytes_RecognizesKnownTypes()
    {
        foreach (var fm in FileMagic.Values)
        {
            if (ReferenceEquals(fm, FileMagic.Unknown))
            {
                continue;
            }

            foreach (var magic in fm.MagicPatterns)
            {
                Assert.Same(fm, FileMagic.ValueOf(magic));
            }
        }

        Assert.Same(FileMagic.Unknown, FileMagic.ValueOf(Encoding.UTF8.GetBytes("foobaa")));
    }
}

