using System.Reflection;
using Xunit;

namespace DotnetPoi.POIFS.Tests;

public class ProjectShellTests
{
    [Fact]
    public void POIFSAssembly_Loads()
    {
        var assembly = Assembly.Load("DotnetPoi.POIFS");

        Assert.Equal("DotnetPoi.POIFS", assembly.GetName().Name);
    }
}
