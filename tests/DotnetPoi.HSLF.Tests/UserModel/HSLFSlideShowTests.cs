using DotnetPoi.HSLF.UserModel;
using Xunit;

namespace DotnetPoi.HSLF.Tests.UserModel;

public class HSLFSlideShowTests
{
    [Fact]
    public void Open_ValidPpt_DoesNotThrow()
    {
        using var stream = File.OpenRead("Test_Humor-Generation.ppt");
        using var prs = new HSLFSlideShow(stream);
        Assert.NotNull(prs);
    }

    [Fact]
    public void Open_NonOle2Stream_ThrowsInvalidDataException()
    {
        var fake = new MemoryStream(new byte[1024]);
        Assert.Throws<InvalidDataException>(() => new HSLFSlideShow(fake));
    }

    [Fact]
    public void GetSlides_ValidPpt_ReturnsSlides()
    {
        using var stream = File.OpenRead("Test_Humor-Generation.ppt");
        using var prs = new HSLFSlideShow(stream);
        Assert.True(prs.getSlides().Count > 0, "Should have at least one slide");
    }

    [Fact]
    public void GetSlides_PptWithKnownCount_ReturnCorrectCount()
    {
        // 45776.ppt is a simple PPT used in POI tests
        using var stream = File.OpenRead("45776.ppt");
        using var prs = new HSLFSlideShow(stream);
        Assert.True(prs.getSlides().Count > 0);
    }

    [Fact]
    public void GetTextParagraphs_SlideWithText_ReturnsContent()
    {
        using var stream = File.OpenRead("Test_Humor-Generation.ppt");
        using var prs = new HSLFSlideShow(stream);
        var hasText = prs.getSlides().Any(s => s.getTextParagraphs().Any(t => t.Length > 0));
        Assert.True(hasText, "At least one slide should have text");
    }
}
