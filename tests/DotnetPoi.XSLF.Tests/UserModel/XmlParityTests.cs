using System.IO.Compression;
using DotnetPoi.XSLF.UserModel;
using Xunit;

namespace DotnetPoi.XSLF.Tests.UserModel;

public class XmlParityTests
{
    [Fact(Skip = "TODO: [dotnet-poi] Align pptx XML output with POI fixtures (Phase 7 Step 1).")]
    public void XmlParity_PptxBasic_MatchesPoiFixtures()
    {
        using var prs = new XMLSlideShow();
        prs.createSlide();

        using var stream = new MemoryStream();
        prs.write(stream);

        CompareToFixtures(stream, "pptx-basic");
    }

    [Fact(Skip = "TODO: [dotnet-poi] Align pptm XML output with POI fixtures (Phase 7 Step 1).")]
    public void XmlParity_PptmBasic_MatchesPoiFixtures()
    {
        using var prs = new XMLSlideShow(new MemoryStream(File.ReadAllBytes("example.pptm")));

        using var stream = new MemoryStream();
        prs.write(stream);

        CompareToFixtures(stream, "pptm-basic");
    }

    private static void CompareToFixtures(MemoryStream stream, string caseName)
    {
        stream.Position = 0;
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read);

        foreach (var fixture in EnumerateFixtures(caseName))
        {
            var entry = archive.GetEntry(fixture.EntryName);
            Assert.NotNull(entry);
            using var reader = new StreamReader(entry.Open());
            var actual = reader.ReadToEnd();
            Assert.Equal(fixture.Expected, actual);
        }
    }

    private static IEnumerable<(string EntryName, string Expected)> EnumerateFixtures(string caseName)
    {
        var fixturesRoot = GetFixturesRoot();
        var prefix = caseName + "__";

        foreach (var file in Directory.EnumerateFiles(fixturesRoot, prefix + "*"))
        {
            var fileName = Path.GetFileName(file);
            var entryName = fileName.Substring(prefix.Length).Replace("__", "/", StringComparison.Ordinal);
            yield return (entryName, File.ReadAllText(file));
        }
    }

    private static string GetFixturesRoot()
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

        return Path.Combine(dir.FullName, "tests", "DotnetPoi.Interop.Tests", "fixtures", "xml-parity");
    }
}
