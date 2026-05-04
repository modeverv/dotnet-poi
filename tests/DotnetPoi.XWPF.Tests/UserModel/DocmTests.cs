using System.IO.Compression;
using DotnetPoi.POIFS.Crypt;
using DotnetPoi.XWPF.UserModel;
using Xunit;

namespace DotnetPoi.XWPF.Tests.UserModel;

/// <summary>Round-trip tests for docm (macro-enabled docx).</summary>
public class DocmTests
{
    private static byte[] LoadDocm() => File.ReadAllBytes("example.docm");

    [Fact]
    public void Read_Docm_DetectsMacros()
    {
        using var stream = new MemoryStream(LoadDocm());
        using var doc = new XWPFDocument(stream);

        Assert.True(doc.HasMacros, "docm should report HasMacros = true");
        Assert.True(doc.isMacroEnabled());
    }

    [Fact]
    public void Read_Docm_HasParagraphs()
    {
        using var stream = new MemoryStream(LoadDocm());
        using var doc = new XWPFDocument(stream);

        Assert.True(doc.getParagraphs().Count > 0, "docm should have paragraphs");
    }

    [Fact]
    public void RoundTrip_Docm_PreservesVbaProjectBin()
    {
        var originalBytes = LoadDocm();
        byte[] originalVba;
        using (var archive = new ZipArchive(new MemoryStream(originalBytes), ZipArchiveMode.Read))
        {
            var entry = archive.GetEntry("word/vbaProject.bin")!;
            using var s = entry.Open();
            using var ms = new MemoryStream();
            s.CopyTo(ms);
            originalVba = ms.ToArray();
        }

        using var inStream = new MemoryStream(originalBytes);
        using var doc = new XWPFDocument(inStream);

        using var outStream = new MemoryStream();
        doc.write(outStream);

        outStream.Position = 0;
        using var archive2 = new ZipArchive(outStream, ZipArchiveMode.Read);
        var vbaEntry = archive2.GetEntry("word/vbaProject.bin");
        Assert.NotNull(vbaEntry);
        using var vs = vbaEntry.Open();
        using var vms = new MemoryStream();
        vs.CopyTo(vms);
        Assert.Equal(originalVba, vms.ToArray());
    }

    [Fact]
    public void RoundTrip_Docm_WritesDocmContentType()
    {
        using var inStream = new MemoryStream(LoadDocm());
        using var doc = new XWPFDocument(inStream);

        using var outStream = new MemoryStream();
        doc.write(outStream);

        outStream.Position = 0;
        using var archive = new ZipArchive(outStream, ZipArchiveMode.Read);
        using var r = new StreamReader(archive.GetEntry("[Content_Types].xml")!.Open());
        var ct = r.ReadToEnd();
        Assert.Contains("macroEnabled.main+xml", ct);
        Assert.Contains("vbaProject", ct);

        using var rr = new StreamReader(archive.GetEntry("word/_rels/document.xml.rels")!.Open());
        var rels = rr.ReadToEnd();
        Assert.Contains("relationships/vbaProject", rels);
        Assert.Contains("Target=\"vbaProject.bin\"", rels);
    }

    [Fact]
    public void RoundTrip_Docm_StillHasMacros()
    {
        using var inStream = new MemoryStream(LoadDocm());
        using var doc = new XWPFDocument(inStream);

        using var outStream = new MemoryStream();
        doc.write(outStream);

        outStream.Position = 0;
        using var doc2 = new XWPFDocument(outStream);
        Assert.True(doc2.HasMacros);
        Assert.Equal(doc.getParagraphs().Count, doc2.getParagraphs().Count);
    }

    [Fact]
    public void Write_DocmFromVbaProject_WritesMacroPackageParts()
    {
        var vbaProject = ReadEntryBytes(LoadDocm(), "word/vbaProject.bin");
        using var doc = new XWPFDocument();
        doc.createParagraph().createRun().setText("macro document");
        doc.setVBAProject(vbaProject);

        using var outStream = new MemoryStream();
        doc.write(outStream);

        outStream.Position = 0;
        using var archive = new ZipArchive(outStream, ZipArchiveMode.Read);
        Assert.Equal(vbaProject, ReadEntryBytes(archive, "word/vbaProject.bin"));
        Assert.Contains("macroEnabled.main+xml", ReadEntry(archive, "[Content_Types].xml"));
        Assert.Contains("relationships/vbaProject", ReadEntry(archive, "word/_rels/document.xml.rels"));
    }

    [Fact]
    public void EncryptDecrypt_Docm_RoundTripsMacros()
    {
        var originalVba = ReadEntryBytes(LoadDocm(), "word/vbaProject.bin");
        using var doc = new XWPFDocument(new MemoryStream(LoadDocm()));
        using var package = new MemoryStream();
        doc.write(package);

        var decrypted = EncryptAndDecryptPackage(package.ToArray(), "f");
        using var decryptedDoc = new XWPFDocument(new MemoryStream(decrypted));

        Assert.True(decryptedDoc.HasMacros);
        Assert.Equal(doc.getParagraphs().Count, decryptedDoc.getParagraphs().Count);
        using var archive = new ZipArchive(new MemoryStream(decrypted), ZipArchiveMode.Read);
        Assert.Equal(originalVba, ReadEntryBytes(archive, "word/vbaProject.bin"));
        Assert.Contains("macroEnabled.main+xml", ReadEntry(archive, "[Content_Types].xml"));
        Assert.Contains("relationships/vbaProject", ReadEntry(archive, "word/_rels/document.xml.rels"));
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
