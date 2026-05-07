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

## Width, Header Rows, and Merged Cells

```csharp
table.setWidth(7200, "dxa");

var header = table.createRow();
header.setHeader(true);
header.setHeight(400, "atLeast");

var merged = header.createCell();
merged.setGridSpan(2);
merged.addParagraph().createRun().setText("Merged header");
```

Vertical merges can be represented with `setVMerge("restart")` on the first cell and `setVMerge("continue")` on following cells.

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

- Grid span, vertical merge, width, row height, header rows, and vertical alignment have modeled APIs.
- Existing borders, shading, cell margins, table layout, and similar unmodeled `tblPr` / `trPr` / `tcPr` children are preserved as raw XML during round-trip.
- API-level creation of detailed border and shading styles is not yet modeled.

## Full Runnable Example

See `examples/UsageSamples/Program.cs` (`CreateDocument`):

[examples/UsageSamples](https://github.com/modeverv/dotnet-poi/tree/main/examples/UsageSamples)
