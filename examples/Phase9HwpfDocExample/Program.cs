using DotnetPoi.HWPF.UserModel;

// This example demonstrates basic text extraction and limited body editing
// for legacy Word 97-2003 (.doc) files using HWPF.

var repoRoot = FindRepositoryRoot(AppContext.BaseDirectory);
var inputPath = Path.Combine(repoRoot, "poi", "test-data", "document", "SampleDoc.doc");
var outputPath = Path.Combine(repoRoot, "examples", "output", "phase9-hwpf-doc-example.doc");

if (!File.Exists(inputPath))
{
    Console.WriteLine($"Input fixture not found: {inputPath}");
    Console.WriteLine("Please ensure the poi submodule is initialized.");
    return;
}

Console.WriteLine($"Reading legacy .doc file: {inputPath}");
using (var stream = File.OpenRead(inputPath))
using (var doc = new HWPFDocument(stream))
{
    // 1. Basic text extraction
    string text = doc.getText();
    Console.WriteLine("--- Extracted Text ---");
    Console.WriteLine(text.Length > 200 ? text.Substring(0, 200) + "..." : text);
    Console.WriteLine("----------------------");

    // 2. Limited body editing
    Console.WriteLine("Appending a paragraph and saving...");
    doc.appendParagraph("\nThis paragraph was added by dotnet-poi HWPF.");
    
    Directory.CreateDirectory(Path.GetDirectoryName(outputPath)!);
    using (var output = File.Create(outputPath))
    {
        doc.write(output);
    }
}

Console.WriteLine($"Saved edited .doc to: {outputPath}");

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
