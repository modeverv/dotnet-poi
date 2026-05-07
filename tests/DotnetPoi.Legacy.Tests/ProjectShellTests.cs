using System.Reflection;
using Xunit;

namespace DotnetPoi.Legacy.Tests;

public class ProjectShellTests
{
    [Fact]
    public void LegacyAssembly_Loads()
    {
        var assembly = Assembly.Load("DotnetPoi.Legacy");

        Assert.Equal("DotnetPoi.Legacy", assembly.GetName().Name);
    }
}
