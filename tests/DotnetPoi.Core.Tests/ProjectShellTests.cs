using Xunit;

namespace DotnetPoi.Core.Tests;

public sealed class ProjectShellTests
{
    [Fact]
    public void CoreCompatibilityAssembly_Loads()
    {
        var assembly = typeof(ProjectShellTests).Assembly;

        Assert.Equal("DotnetPoi.Core.Tests", assembly.GetName().Name);
    }
}
