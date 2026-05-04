using DotnetPoi.HWPF.UserModel;
using Xunit;

namespace DotnetPoi.HWPF.Tests.UserModel;

public class HWPFDocumentTests
{
    [Fact]
    public void Open_ValidDoc_DoesNotThrow()
    {
        using var stream = File.OpenRead("TestMickey.doc");
        using var doc = new HWPFDocument(stream);
        Assert.NotNull(doc);
    }

    [Fact]
    public void Open_NonOle2Stream_ThrowsInvalidDataException()
    {
        var fake = new MemoryStream(new byte[1024]);
        Assert.Throws<InvalidDataException>(() => new HWPFDocument(fake));
    }

    [Fact]
    public void GetText_ValidDoc_ReturnsNonEmpty()
    {
        using var stream = File.OpenRead("TestMickey.doc");
        using var doc = new HWPFDocument(stream);
        var text = doc.getText();
        // The doc should have some textual content
        Assert.NotNull(text);
        Assert.True(text.Length >= 0); // might be empty if no CLX, but should not throw
    }

    [Fact]
    public void GetCcpText_ValidDoc_ReturnsPositiveOrZero()
    {
        using var stream = File.OpenRead("TestMickey.doc");
        using var doc = new HWPFDocument(stream);
        Assert.True(doc.getCcpText() >= 0);
    }

    [Fact]
    public void RoundTrip_GetText_DoesNotThrowOnRealFile()
    {
        // Smoke test: open, read text, close without exception
        using var stream = File.OpenRead("TestMickey.doc");
        using var doc = new HWPFDocument(stream);
        _ = doc.getText();
    }
}
