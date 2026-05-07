# Rich Text

Apply different formatting to individual characters within a single cell using `XSSFRichTextString`.

## Basic Usage

```csharp
using DotnetPoi.XSSF.UserModel;

var rich = new XSSFRichTextString();
rich.addRun("Bold prefix ", bold: true);
rich.addRun("and italic suffix.", bold: false, italic: true);

var row = sheet.createRow(0);
var cell = row.createCell(0);
cell.setCellValue(rich);
```

## Per-Character Formatting

The `addRun` method applies formatting to a substring:

```csharp
rich.addRun(
    text: "formatted text",
    bold: false,
    italic: true,
    underline: false,
    strikethrough: false,
    fontName: "Arial",
    fontSize: 14,
    color: "FF0000"
);
```

## Reading Rich Text

```csharp
var rich = cell.getRichStringCellValue();
var fullText = rich.getString();  // plain text concatenation

// Access formatting runs
var numRuns = rich.numFormattingRuns();
for (int i = 0; i < numRuns; i++)
{
    var start = rich.getIndexOfFormattingRun(i);
    var end = rich.getEndIndexOfFormattingRun(i);
    var text = fullText.Substring(start, end - start);
    // Check formatting properties...
}
```

## Full Runnable Example

See `examples/UsageSamples/Program.cs` (`CreateSpreadsheet`):

[examples/UsageSamples](https://github.com/modeverv/dotnet-poi/tree/master/examples/UsageSamples)
