namespace DotnetPoi.SS.Tests.Xml;

internal static class XmlFixturePaths
{
    public static string GetRepositoryRoot()
    {
        var dir = new DirectoryInfo(AppContext.BaseDirectory);
        while (dir != null && !File.Exists(Path.Combine(dir.FullName, "DotnetPOI.sln")))
        {
            dir = dir.Parent;
        }

        if (dir == null)
        {
            throw new DirectoryNotFoundException("Could not locate repository root (DotnetPOI.sln).");
        }

        return dir.FullName;
    }

    public static string GetFixturePath(string fileName)
    {
        return Path.Combine(GetRepositoryRoot(),
            "tests",
            "DotnetPoi.Interop.Tests",
            "fixtures",
            "xml-parity",
            fileName);
    }
}
