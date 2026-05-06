# Tables

Create and read tables in docx documents.

## Creating a Table

```csharp
using DotnetPoi.XWPF.UserModel;

using var doc = new XWPFDocument();
var table = doc.createTable();

// Set column widths
table.addGridCol(3600);
table.addGridCol(3600);

// Add rows
var row1 = table.createRow();
row1.createCell().addParagraph().createRun().setText("Name");
row1.createCell().addParagraph().createRun().setText("Value");

var row2 = table.createRow();
row2.createCell().addParagraph().createRun().setText("Item 1");
row2.createCell().addParagraph().createRun().setText("42");
```

## Reading Tables

```csharp
using var stream = File.OpenRead("input.docx");
using var doc = new XWPFDocument(stream);

var tables = doc.getTables();
var table = tables[0];
var rows = table.getRows();
foreach (var row in rows)
{
    var cells = row.getCells();
    foreach (var cell in cells)
    {
        Console.WriteLine(cell.getParagraphs()[0].getText());
    }
}
```

## Limitations

- Cell merging is not supported
- Table borders are not supported
- Cell-level formatting (shading, alignment) is not supported

## Full Runnable Example

See `examples/UsageSamples/Program.cs` (`CreateDocument`):

[examples/UsageSamples](https://github.com/modeverv/dotnet-poi/tree/main/examples/UsageSamples)
