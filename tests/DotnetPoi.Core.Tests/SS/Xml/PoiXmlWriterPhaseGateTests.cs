using Xunit;

namespace DotnetPoi.SS.Tests.Xml;

public class PoiXmlWriterPhaseGateTests
{
    private static readonly string[] BannedDirectXmlApis =
    [
        "System.Xml.XmlWriter",
        "XmlWriter.Create",
        "new XmlWriterSettings",
        "System.Xml.Linq",
        "XDocument",
        "XElement",
        "XmlDocument",
        "XmlSerializer"
    ];

    [Fact]
    public void Source_OoxmlXmlOutput_DoesNotBypassPoiXmlWriter()
    {
        var repoRoot = XmlFixturePaths.GetRepositoryRoot();
        var sourceRoot = Path.Combine(repoRoot, "src");
        var violations = new List<string>();

        foreach (var file in Directory.EnumerateFiles(sourceRoot, "*.cs", SearchOption.AllDirectories))
        {
            if (IsGeneratedOrBuildOutput(file))
            {
                continue;
            }

            if (Path.GetFileName(file) == "PoiXmlWriter.cs")
            {
                continue;
            }

            var text = File.ReadAllText(file);
            foreach (var bannedApi in BannedDirectXmlApis)
            {
                if (text.Contains(bannedApi, StringComparison.Ordinal))
                {
                    violations.Add($"{Path.GetRelativePath(repoRoot, file)} uses {bannedApi}");
                }
            }
        }

        Assert.True(violations.Count == 0,
            "OOXML XML output must go through PoiXmlWriter. Direct XML APIs found:" +
            Environment.NewLine +
            string.Join(Environment.NewLine, violations));
    }

    private static bool IsGeneratedOrBuildOutput(string file)
    {
        var normalized = file.Replace(Path.DirectorySeparatorChar, '/');
        return normalized.Contains("/bin/", StringComparison.Ordinal) ||
            normalized.Contains("/obj/", StringComparison.Ordinal) ||
            normalized.EndsWith(".g.cs", StringComparison.Ordinal);
    }
}
