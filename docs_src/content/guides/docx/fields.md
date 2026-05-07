# Fields

dotnet-poi supports fields including Table of Contents (TOC), page numbers, and mail merge fields.

## Page Number Field

```csharp
using DotnetPoi.XWPF.UserModel;

using var doc = new XWPFDocument();

var p = doc.createParagraph();
var r = p.createRun();
r.setText("Page ");
p.addField(" PAGE ", "1");
```

## TOC (Table of Contents)

```csharp
var tocParagraph = doc.createParagraph();
tocParagraph.addField("TOC \\o \"1-3\" \\h \\z \\u");
```

## Mail Merge Fields

```csharp
paragraph.addField("MERGEFIELD CustomerName", "Acme Inc.");
```

## Reading Fields

```csharp
var fields = paragraph.getFields();
foreach (var f in fields)
{
    var code = f.Instruction;
    // "PAGE", "TOC \\o \"1-3\" \\h \\z \\u", "MERGEFIELD CustomerName"
}
```

Fields are preserved on round-trip and are interoperable with Java POI.

## Full Runnable Example

See `examples/UsageSamples/Program.cs` (`CreateDocument`):

[examples/UsageSamples](https://github.com/modeverv/dotnet-poi/tree/main/examples/UsageSamples)
