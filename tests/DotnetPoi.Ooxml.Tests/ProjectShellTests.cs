using System.Reflection;
using Xunit;

namespace DotnetPoi.Ooxml.Tests;

public class ProjectShellTests
{
    [Fact]
    public void OoxmlAssembly_Loads()
    {
        var assembly = Assembly.Load("DotnetPoi.Ooxml");

        Assert.Equal("DotnetPoi.Ooxml", assembly.GetName().Name);
    }
}
