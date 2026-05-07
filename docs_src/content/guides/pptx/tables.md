# Tables

Create tables on pptx slides.

## Creating a Table

```csharp
using DotnetPoi.XSLF.UserModel;

using var ppt = new XMLSlideShow();
var slide = ppt.createSlide();

// Create table and set position
var table = slide.createTable();
table.setAnchor(5715000, 914400, 2743200, 1371600);

// Define column widths
table.addGridCol(1371600);
table.addGridCol(1371600);
```

## Adding Rows and Cells

```csharp
// Add a row with cells
var row1 = table.createRow();
row1.createCell().addParagraph().addRun("Header 1");
row1.createCell().addParagraph().addRun("Header 2");

var row2 = table.createRow();
row2.createCell().addParagraph().addRun("Value A");
row2.createCell().addParagraph().addRun("Value B");
```

## Reading Tables

```csharp
var tables = slide.getTables();
var table = tables[0];
foreach (var row in table.Rows)
{
    foreach (var cell in row.Cells)
    {
        var text = cell.Paragraphs[0].getPlainText();
        Console.Write($"{text}\t");
    }
    Console.WriteLine();
}
```

## Full Runnable Example

See `examples/Phase33PptxExample/` and `examples/UsageSamples/Program.cs` (`CreatePresentation`):

[examples/Phase33PptxExample](https://github.com/modeverv/dotnet-poi/tree/master/examples/Phase33PptxExample)

[examples/UsageSamples](https://github.com/modeverv/dotnet-poi/tree/master/examples/UsageSamples)
