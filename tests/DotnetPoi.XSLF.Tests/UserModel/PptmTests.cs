using System.IO.Compression;
using DotnetPoi.POIFS.Crypt;
using DotnetPoi.XSLF.UserModel;
using Xunit;

namespace DotnetPoi.XSLF.Tests.UserModel;

/// <summary>Round-trip tests for pptm (macro-enabled pptx).</summary>
public class PptmTests
{
    private static byte[] LoadPptm() => File.ReadAllBytes("example.pptm");

    [Fact]
    public void Read_Pptm_DetectsMacros()
    {
        using var stream = new MemoryStream(LoadPptm());
        using var prs = new XMLSlideShow(stream);

        Assert.True(prs.HasMacros, "pptm should report HasMacros = true");
        Assert.True(prs.isMacroEnabled());
    }

    [Fact]
    public void Read_Pptm_HasSlides()
    {
        using var stream = new MemoryStream(LoadPptm());
        using var prs = new XMLSlideShow(stream);

        Assert.True(prs.getSlides().Count > 0, "pptm should have at least one slide");
    }

    [Fact]
    public void RoundTrip_Pptm_PreservesVbaProjectBin()
    {
        var originalBytes = LoadPptm();
        byte[] originalVba;
        using (var archive = new ZipArchive(new MemoryStream(originalBytes), ZipArchiveMode.Read))
        {
            var entry = archive.GetEntry("ppt/vbaProject.bin")!;
            using var s = entry.Open();
            using var ms = new MemoryStream();
            s.CopyTo(ms);
            originalVba = ms.ToArray();
        }

        using var inStream = new MemoryStream(originalBytes);
        using var prs = new XMLSlideShow(inStream);

        using var outStream = new MemoryStream();
        prs.write(outStream);

        outStream.Position = 0;
        using var archive2 = new ZipArchive(outStream, ZipArchiveMode.Read);
        var vbaEntry = archive2.GetEntry("ppt/vbaProject.bin");
        Assert.NotNull(vbaEntry);
        using var vs = vbaEntry.Open();
        using var vms = new MemoryStream();
        vs.CopyTo(vms);
        Assert.Equal(originalVba, vms.ToArray());
    }

    [Fact]
    public void RoundTrip_Pptm_WritesPptmContentType()
    {
        using var inStream = new MemoryStream(LoadPptm());
        using var prs = new XMLSlideShow(inStream);

        using var outStream = new MemoryStream();
        prs.write(outStream);

        outStream.Position = 0;
        using var archive = new ZipArchive(outStream, ZipArchiveMode.Read);
        using var r = new StreamReader(archive.GetEntry("[Content_Types].xml")!.Open());
        var ct = r.ReadToEnd();
        Assert.Contains("macroEnabled.main+xml", ct);
        Assert.Contains("vbaProject", ct);

        using var rr = new StreamReader(archive.GetEntry("ppt/_rels/presentation.xml.rels")!.Open());
        var rels = rr.ReadToEnd();
        Assert.Contains("relationships/vbaProject", rels);
        Assert.Contains("Target=\"vbaProject.bin\"", rels);
    }

    [Fact]
    public void RoundTrip_Pptm_StillHasMacros()
    {
        using var inStream = new MemoryStream(LoadPptm());
        using var prs = new XMLSlideShow(inStream);
        var slideCount = prs.getSlides().Count;

        using var outStream = new MemoryStream();
        prs.write(outStream);

        outStream.Position = 0;
        using var prs2 = new XMLSlideShow(outStream);
        Assert.True(prs2.HasMacros);
        Assert.Equal(slideCount, prs2.getSlides().Count);
    }

    [Fact]
    public void Write_PptmFromVbaProject_WritesMacroPackageParts()
    {
        var vbaProject = ReadEntryBytes(LoadPptm(), "ppt/vbaProject.bin");
        using var prs = new XMLSlideShow();
        prs.createSlide();
        prs.setVBAProject(vbaProject);

        using var outStream = new MemoryStream();
        prs.write(outStream);

        outStream.Position = 0;
        using var archive = new ZipArchive(outStream, ZipArchiveMode.Read);
        Assert.Equal(vbaProject, ReadEntryBytes(archive, "ppt/vbaProject.bin"));
        Assert.Contains("macroEnabled.main+xml", ReadEntry(archive, "[Content_Types].xml"));
        Assert.Contains("relationships/vbaProject", ReadEntry(archive, "ppt/_rels/presentation.xml.rels"));
    }

    [Fact]
    public void EncryptDecrypt_Pptm_RoundTripsMacros()
    {
        var originalVba = ReadEntryBytes(LoadPptm(), "ppt/vbaProject.bin");
        using var prs = new XMLSlideShow(new MemoryStream(LoadPptm()));
        var slideCount = prs.getSlides().Count;
        using var package = new MemoryStream();
        prs.write(package);

        var decrypted = EncryptAndDecryptPackage(package.ToArray(), "f");
        using var decryptedPrs = new XMLSlideShow(new MemoryStream(decrypted));

        Assert.True(decryptedPrs.HasMacros);
        Assert.Equal(slideCount, decryptedPrs.getSlides().Count);
        using var archive = new ZipArchive(new MemoryStream(decrypted), ZipArchiveMode.Read);
        Assert.Equal(originalVba, ReadEntryBytes(archive, "ppt/vbaProject.bin"));
        Assert.Contains("macroEnabled.main+xml", ReadEntry(archive, "[Content_Types].xml"));
        Assert.Contains("relationships/vbaProject", ReadEntry(archive, "ppt/_rels/presentation.xml.rels"));
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

    private static byte[] EncryptAndDecryptPackage(byte[] packageBytes, string password)
    {
        var info = new EncryptionInfo(EncryptionMode.agile);
        info.Encryptor.confirmPassword(password);
        using var encrypted = new MemoryStream();
        info.Encryptor.encryptPackage(packageBytes, encrypted);

        encrypted.Position = 0;
        var read = new EncryptionInfo(encrypted);
        Assert.True(read.Decryptor.verifyPassword(password));
        return read.Decryptor.getData();
    }
}
