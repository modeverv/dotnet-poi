
using DotnetPoi.XSLF.UserModel;

var repoRoot = FindRepositoryRoot(AppContext.BaseDirectory);
var outputDirectory = Path.Combine(repoRoot, "examples", "output");
Directory.CreateDirectory(outputDirectory);

var outputPath = Path.Combine(outputDirectory, "phase3_3-pptx-example.pptx");
var imageBytes = File.ReadAllBytes(Path.Combine(repoRoot, "tests", "image.jpg"));

WritePresentation(outputPath, imageBytes);
ReadAndAssertPresentation(outputPath);

Console.WriteLine("Phase 3.3 pptx example passed.");
Console.WriteLine($"dotnet-poi wrote: {outputPath}");

static void WritePresentation(string outputPath, byte[] imageBytes)
{
    using var prs = new XMLSlideShow();

    var picIdx = prs.addPicture(imageBytes, XSLFPictureData.PICTURE_TYPE_JPEG);

    // Slide 1: photo filling the entire slide at 0° rotation
    var slide1 = prs.createSlide();
    var shape1 = prs.createPicture(slide1, picIdx);
    shape1.setAnchor(0, 0, XMLSlideShow.DefaultSlideCx, XMLSlideShow.DefaultSlideCy);

    // Slide 2: same photo, placed at 1 inch offset, 3×2.25 inch size, 45° rotation, horizontal flip
    var slide2 = prs.createSlide();
    var shape2 = prs.createPicture(slide2, picIdx);
    shape2.setAnchor(914_400, 685_800, 2_743_200, 2_057_400);
    shape2.setRotation(45.0);
    shape2.setFlipHorizontal(true);

    // Slide 3: same photo at 1 inch × 1 inch, rotated 90°
    var slide3 = prs.createSlide();
    var shape3 = prs.createPicture(slide3, picIdx);
    shape3.setAnchor(914_400, 685_800, 2_743_200, 2_057_400);
    shape3.setRotation(90.0);

    using var stream = File.Create(outputPath);
    prs.write(stream);
}

static void ReadAndAssertPresentation(string path)
{
    using var stream = File.OpenRead(path);
    using var prs    = new XMLSlideShow(stream);

    AssertInt(3, prs.getSlides().Count, "slide count");
    // All 3 slides share the same JPEG picture data (deduplicated)
    AssertInt(1, prs.getPictureData().Count, "picture data count");
    AssertInt(XSLFPictureData.PICTURE_TYPE_JPEG, prs.getPictureData()[0].getPictureType(), "picture type");

    // Slide 1 — full slide, no rotation
    var slide1 = prs.getSlides()[0];
    AssertInt(1, slide1.getShapes().Count, "slide 1 shape count");
    var s1 = slide1.getShapes()[0];
    AssertDouble(0.0, s1.getRotation(), "slide 1 rotation");
    AssertLong(0L, s1.getAnchorX(), "slide 1 anchor x");
    AssertLong(0L, s1.getAnchorY(), "slide 1 anchor y");
    AssertLong(XMLSlideShow.DefaultSlideCx, s1.getAnchorCx(), "slide 1 width");

    // Slide 2 — 45° rotation, horizontal flip
    var slide2 = prs.getSlides()[1];
    AssertInt(1, slide2.getShapes().Count, "slide 2 shape count");
    var s2 = slide2.getShapes()[0];
    AssertDouble(45.0, s2.getRotation(), "slide 2 rotation");
    if (!s2.getFlipHorizontal())
        throw new InvalidOperationException("slide 2 shape should have flipH=true");
    AssertLong(914_400L,   s2.getAnchorX(),  "slide 2 anchor x");
    AssertLong(2_743_200L, s2.getAnchorCx(), "slide 2 width");

    // Slide 3 — 90° rotation
    var slide3 = prs.getSlides()[2];
    AssertInt(1, slide3.getShapes().Count, "slide 3 shape count");
    var s3 = slide3.getShapes()[0];
    AssertDouble(90.0, s3.getRotation(), "slide 3 rotation");
}

static void AssertInt(int expected, int actual, string label)
{
    if (expected != actual)
        throw new InvalidOperationException($"{label}: expected {expected}, got {actual}.");
}

static void AssertLong(long expected, long actual, string label)
{
    if (expected != actual)
        throw new InvalidOperationException($"{label}: expected {expected}, got {actual}.");
}

static void AssertDouble(double expected, double actual, string label)
{
    if (Math.Abs(expected - actual) > 1e-6)
        throw new InvalidOperationException($"{label}: expected {expected}, got {actual}.");
}

static string FindRepositoryRoot(string startDirectory)
{
    var directory = new DirectoryInfo(startDirectory);
    while (directory is not null)
    {
        if (File.Exists(Path.Combine(directory.FullName, "DotnetPOI.sln")))
            return directory.FullName;
        directory = directory.Parent;
    }
    throw new DirectoryNotFoundException("Could not locate the dotnet-poi repository root.");
}
