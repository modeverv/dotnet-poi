using System.IO.Compression;
using DotnetPoi.POIFS.Crypt;
using DotnetPoi.SS.UserModel;
using DotnetPoi.XSSF.UserModel;
using Xunit;

namespace DotnetPoi.XSSF.Tests.UserModel;

/// <summary>
/// Round-trip tests for xlsm (macro-enabled xlsx).
/// Ported approach: vbaProject.bin is carried through byte-for-byte;
/// cell data is readable via the normal XSSF API.
/// </summary>
public class XlsmTests
{
    private static byte[] LoadXlsm() => File.ReadAllBytes("example.xlsm");

    [Fact]
    public void Read_Xlsm_LoadsSheets()
    {
        using var stream = new MemoryStream(LoadXlsm());
        using var wb = new XSSFWorkbook(stream);

        Assert.True(wb.getNumberOfSheets() > 0, "xlsm should have at least one sheet");
    }

    [Fact]
    public void Read_Xlsm_ContainsVbaProjectBin()
    {
        using var stream = new MemoryStream(LoadXlsm());
        using var wb = new XSSFWorkbook(stream);

        Assert.True(wb.HasMacros, "xlsm should report HasMacros = true");
        Assert.True(wb.isMacroEnabled());
    }

    [Fact]
    public void RoundTrip_Xlsm_PreservesVbaProjectBin()
    {
        var originalBytes = LoadXlsm();
        byte[] originalVba;
        using (var archive = new ZipArchive(new MemoryStream(originalBytes), ZipArchiveMode.Read))
        {
            var entry = archive.GetEntry("xl/vbaProject.bin")!;
            using var s = entry.Open();
            using var ms = new MemoryStream();
            s.CopyTo(ms);
            originalVba = ms.ToArray();
        }

        // Read → write → read back
        using var inStream = new MemoryStream(originalBytes);
        using var wb = new XSSFWorkbook(inStream);

        using var outStream = new MemoryStream();
        wb.write(outStream);

        outStream.Position = 0;
        using var archive2 = new ZipArchive(outStream, ZipArchiveMode.Read);
        var vbaEntry = archive2.GetEntry("xl/vbaProject.bin");
        Assert.NotNull(vbaEntry);
        using var vs = vbaEntry.Open();
        using var vms = new MemoryStream();
        vs.CopyTo(vms);
        Assert.Equal(originalVba, vms.ToArray());
    }

    [Fact]
    public void RoundTrip_Xlsm_WritesXlsmContentType()
    {
        using var inStream = new MemoryStream(LoadXlsm());
        using var wb = new XSSFWorkbook(inStream);

        using var outStream = new MemoryStream();
        wb.write(outStream);

        outStream.Position = 0;
        using var archive = new ZipArchive(outStream, ZipArchiveMode.Read);
        var ct = ReadEntry(archive, "[Content_Types].xml");
        Assert.Contains("macroEnabled.main+xml", ct);
        Assert.Contains("vbaProject", ct);

        var rels = ReadEntry(archive, "xl/_rels/workbook.xml.rels");
        Assert.Contains("relationships/vbaProject", rels);
        Assert.Contains("Target=\"vbaProject.bin\"", rels);
    }

    [Fact]
    public void RoundTrip_Xlsm_CellDataPreserved()
    {
        using var inStream = new MemoryStream(LoadXlsm());
        using var wb = new XSSFWorkbook(inStream);

        using var outStream = new MemoryStream();
        wb.write(outStream);

        outStream.Position = 0;
        using var wb2 = new XSSFWorkbook(outStream);
        Assert.Equal(wb.getNumberOfSheets(), wb2.getNumberOfSheets());
        Assert.True(wb2.HasMacros);
    }

    [Fact]
    public void Write_XlsmFromVbaProject_WritesMacroPackageParts()
    {
        var vbaProject = ReadEntryBytes(LoadXlsm(), "xl/vbaProject.bin");
        using var wb = new XSSFWorkbook();
        wb.createSheet("MacroSheet");
        wb.setVBAProject(vbaProject);

        using var outStream = new MemoryStream();
        wb.write(outStream);

        outStream.Position = 0;
        using var archive = new ZipArchive(outStream, ZipArchiveMode.Read);
        Assert.Equal(vbaProject, ReadEntryBytes(archive, "xl/vbaProject.bin"));
        Assert.Contains("macroEnabled.main+xml", ReadEntry(archive, "[Content_Types].xml"));
        Assert.Contains("relationships/vbaProject", ReadEntry(archive, "xl/_rels/workbook.xml.rels"));
    }

    [Fact]
    public void WriteEncrypted_Xlsm_RoundTripsMacrosAfterDecrypt()
    {
        var originalVba = ReadEntryBytes(LoadXlsm(), "xl/vbaProject.bin");
        using var wb = new XSSFWorkbook(new MemoryStream(LoadXlsm()));

        using var encrypted = new MemoryStream();
        wb.writeEncrypted(encrypted, "f");

        var decrypted = DecryptPackage(encrypted.ToArray(), "f");
        using var decryptedWorkbook = new XSSFWorkbook(new MemoryStream(decrypted));

        Assert.True(decryptedWorkbook.HasMacros);
        Assert.Equal(wb.getNumberOfSheets(), decryptedWorkbook.getNumberOfSheets());
        using var archive = new ZipArchive(new MemoryStream(decrypted), ZipArchiveMode.Read);
        Assert.Equal(originalVba, ReadEntryBytes(archive, "xl/vbaProject.bin"));
        Assert.Contains("macroEnabled.main+xml", ReadEntry(archive, "[Content_Types].xml"));
        Assert.Contains("relationships/vbaProject", ReadEntry(archive, "xl/_rels/workbook.xml.rels"));
    }

    private static string ReadEntry(ZipArchive archive, string name)
    {
        var entry = archive.GetEntry(name);
        Assert.NotNull(entry);
        using var r = new StreamReader(entry.Open());
        return r.ReadToEnd();
    }

    private static byte[] ReadEntryBytes(byte[] packageBytes, string name)
    {
        using var archive = new ZipArchive(new MemoryStream(packageBytes), ZipArchiveMode.Read);
        return ReadEntryBytes(archive, name);
    }

    private static byte[] ReadEntryBytes(ZipArchive archive, string name)
    {
        var entry = archive.GetEntry(name);
        Assert.NotNull(entry);
        using var s = entry.Open();
        using var ms = new MemoryStream();
        s.CopyTo(ms);
        return ms.ToArray();
    }

    private static byte[] DecryptPackage(byte[] encryptedPackage, string password)
    {
        using var stream = new MemoryStream(encryptedPackage);
        var info = new EncryptionInfo(stream);
        Assert.True(info.Decryptor.verifyPassword(password));
        return info.Decryptor.getData();
    }
}
