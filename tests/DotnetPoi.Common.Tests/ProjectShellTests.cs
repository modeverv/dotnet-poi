using System.Reflection;
using Xunit;

namespace DotnetPoi.Common.Tests;

public class ProjectShellTests
{
    [Fact]
    public void CommonAssembly_Loads()
    {
        var assembly = Assembly.Load("DotnetPoi.Common");

        Assert.Equal("DotnetPoi.Common", assembly.GetName().Name);
    }
}
