# Fields

dotnet-poi supports fields including Table of Contents (TOC), page numbers, and mail merge fields.

## Page Number Field

```csharp
using DotnetPoi.XWPF.UserModel;

using var doc = new XWPFDocument();

var p = doc.createParagraph();
var r = p.createRun();
r.setText("Page ");

var field = p.addField();
field.SetFieldCode("PAGE");
```

## TOC (Table of Contents)

```csharp
var tocParagraph = doc.createParagraph();
var tocField = tocParagraph.addField();
tocField.SetFieldCode("TOC \\o \"1-3\" \\h \\z \\u");
```

## Mail Merge Fields

```csharp
var mergeField = paragraph.addField();
mergeField.SetFieldCode("MERGEFIELD CustomerName");
```

## Reading Fields

```csharp
var fields = paragraph.getFields();
foreach (var f in fields)
{
    var code = f.GetFieldCode();
    // "PAGE", "TOC \\o \"1-3\" \\h \\z \\u", "MERGEFIELD CustomerName"
}
```

Fields are preserved on round-trip and are interoperable with Java POI.
