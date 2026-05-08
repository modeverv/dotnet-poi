using DotnetPoi.HSLF.UserModel;

// This example demonstrates basic text extraction from legacy
// PowerPoint 97-2003 (.ppt) files using the early HSLF reader.

var repoRoot = FindRepositoryRoot(AppContext.BaseDirectory);
var inputPath = Path.Combine(repoRoot, "poi", "test-data", "slideshow", "SampleShow.ppt");

if (!File.Exists(inputPath))
{
    Console.WriteLine($"Input fixture not found: {inputPath}");
    Console.WriteLine("Please ensure the poi submodule is initialized.");
    return;
}

Console.WriteLine($"Reading legacy .ppt file: {inputPath}");
using (var stream = File.OpenRead(inputPath))
using (var ppt = new HSLFSlideShow(stream))
{
    var slides = ppt.getSlides();
    Console.WriteLine($"Detected {slides.Count} slides.");

    for (int i = 0; i < slides.Count; i++)
    {
        Console.WriteLine($"--- Slide {i + 1} ---");
        foreach (var text in slides[i].getTextParagraphs())
        {
            Console.WriteLine(text);
        }
    }
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
