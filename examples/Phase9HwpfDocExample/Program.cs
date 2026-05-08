using DotnetPoi.HWPF.UserModel;

// This example demonstrates basic text extraction and limited body editing
// for legacy Word 97-2003 (.doc) files using HWPF.

var repoRoot = FindRepositoryRoot(AppContext.BaseDirectory);
var inputPath = Path.Combine(repoRoot, "poi", "test-data", "document", "SampleDoc.doc");
var headerInputPath = Path.Combine(repoRoot, "poi", "test-data", "document", "HeaderFooterUnicode.doc");
var tableInputPath = Path.Combine(repoRoot, "poi", "test-data", "document", "innertable.doc");
var outputPath = Path.Combine(repoRoot, "examples", "output", "phase9-hwpf-doc-example.doc");

Console.WriteLine($"Reading HeaderFooterUnicode.doc file: {headerInputPath}");
using (var stream = File.OpenRead(headerInputPath))
using (var doc = new HWPFDocument(stream))
{
    Console.WriteLine("--- Header Story Text ---");
    Console.WriteLine(doc.getHeaderStoryRange().text());
}

Console.WriteLine($"Reading innertable.doc file: {tableInputPath}");
using (var stream = File.OpenRead(tableInputPath))
using (var doc = new HWPFDocument(stream))
{
    Console.WriteLine("--- Tables ---");
    var range = doc.getRange();
    for (int i = 0; i < range.numParagraphs(); i++)
    {
        var p = range.getParagraph(i);
        if (p.isInTable())
        {
            var table = range.getTable(p);
            Console.WriteLine($"Table Level: {table.getTableLevel()}, Rows: {table.numRows()}");
            for (int r = 0; r < table.numRows(); r++)
            {
                var row = table.getRow(r);
                Console.WriteLine($"  Row {r}, Cells: {row.numCells()}");
                for (int c = 0; c < row.numCells(); c++)
                {
                    Console.WriteLine($"    Cell {c}: {row.getCell(c).text().Replace("\r", "\\r").Replace("\a", "\\a")}");
                }
            }
            // Skip paragraphs in this table
            i += table.numParagraphs() - 1;
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
