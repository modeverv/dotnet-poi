using System.IO.Compression;
using System.Linq;
using DotnetPoi.XSSF.UserModel;
using Xunit;

namespace DotnetPoi.XSSF.Tests.UserModel;

public class XmlParityTests
{
    [Fact]
    public void XmlParity_XlsmBasic_MatchesPoiFixtures()
    {
        var repoRoot = GetRepoRoot();
        var sourcePath = Path.Combine(repoRoot, "tests", "test-files", "example.xlsm");
        using var wb = new XSSFWorkbook(new MemoryStream(File.ReadAllBytes(sourcePath)));

        // Mark workbook as dirty to force XML regeneration instead of preserving original package.
        // This is necessary because the test compares against POI's output, not the original file.
        wb.setForceFormulaRecalculation(wb.getForceFormulaRecalculation());

        using var stream = new MemoryStream();
        wb.write(stream);

        CompareToFixtures(stream, "xlsm-basic");
    }

    private static void CompareToFixtures(MemoryStream stream, string caseName)
    {
        stream.Position = 0;
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read);

        var fixtureFiles = EnumerateFixtures(caseName).ToList();
        Assert.True(fixtureFiles.Any(), $"No fixtures found for case '{caseName}'.");

        foreach (var fixture in fixtureFiles)
        {
            var entry = archive.GetEntry(fixture.EntryName);
            Assert.True(entry != null, $"Entry '{fixture.EntryName}' not found in generated archive.");
            using var reader = new StreamReader(entry.Open());
            var actual = reader.ReadToEnd();

            // Normalize line endings and compare
            var expectedNormalized = fixture.Expected.Replace("\r\n", "\n");
            var actualNormalized = actual.Replace("\r\n", "\n");

            Assert.Equal(expectedNormalized, actualNormalized);
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

    private static string GetRepoRoot()
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
}
