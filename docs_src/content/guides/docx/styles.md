# Styles

Apply and inspect paragraph style references in docx documents.

## Applying a Paragraph Style

```csharp
using DotnetPoi.XWPF.UserModel;

using var doc = new XWPFDocument();
var heading = doc.createParagraph();
heading.setStyle("Heading1");
heading.createRun().setText("Quarterly report");

var body = doc.createParagraph();
body.setStyle("Normal");
body.createRun().setText("Generated with dotnet-poi.");
```

The style ID is written as `w:pStyle` and is preserved when the document is read back and saved again.

## Reading Style References

```csharp
using var stream = File.OpenRead("report.docx");
using var doc = new XWPFDocument(stream);

foreach (var paragraph in doc.getParagraphs())
{
    Console.WriteLine(paragraph.getStyleID() ?? "(direct formatting)");
}
```

## Inspecting styles.xml

```csharp
var styles = doc.getStyles();
var heading1 = styles?.getStyle("Heading1");
Console.WriteLine(heading1?.Name);
```

`XWPFStyles` reads style IDs, names, types, default flags, and `basedOn` relationships from `word/styles.xml`.

## Default Styles

New documents write a default `word/styles.xml` with Normal and Heading 1-3 definitions, following the Apache POI-style baseline used by this port. Existing `word/styles.xml` parts are preserved during round-trip.

## Limitations

- Paragraph style references are modeled with `setStyle()` and `getStyleID()`.
- Character styles and table styles are not yet editable through a public API.
- Full Word style inheritance is not evaluated by dotnet-poi; Word or LibreOffice applies styles when opening the document.

## Full Runnable Example

See `examples/Phase32DocxExample/` for docx creation patterns:

[examples/Phase32DocxExample](https://github.com/modeverv/dotnet-poi/tree/master/examples/Phase32DocxExample)
