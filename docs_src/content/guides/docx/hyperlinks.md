# Hyperlinks

Add external hyperlinks to docx paragraphs.

## Adding a Hyperlink

```csharp
using DotnetPoi.XWPF.UserModel;

using var doc = new XWPFDocument();

var p = doc.createParagraph();
var r = p.createRun();
r.setText("Visit ");
r.setColor("0563C1");
r.setUnderline(UnderlinePatterns.Single);

// Add the hyperlink
var linkRun = p.createRun();
linkRun.setText("our website");
linkRun.setColor("0563C1");
linkRun.setUnderline(UnderlinePatterns.Single);
linkRun.setHyperlink("https://github.com/modeverv/dotnet-poi");
```

Hyperlinks round-trip correctly and are readable by Java POI.

## Full Runnable Example

See `examples/UsageSamples/Program.cs` (`CreateDocument`):

[examples/UsageSamples](https://github.com/modeverv/dotnet-poi/tree/master/examples/UsageSamples)
