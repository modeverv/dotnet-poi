# Headers and Footers

Add headers and footers to docx documents.

## Adding a Header

```csharp
using DotnetPoi.XWPF.UserModel;

using var doc = new XWPFDocument();

var header = doc.getHeaderFooterPolicy().createHeader(XWPFHeaderFooterPolicy.DEFAULT);
var hPara = header.createParagraph();
var hRun = hPara.createRun();
hRun.setText("Document Header");
```

## Adding a Footer

```csharp
var footer = doc.getHeaderFooterPolicy().createFooter(XWPFHeaderFooterPolicy.DEFAULT);
var fPara = footer.createParagraph();
var fRun = fPara.createRun();
fRun.setText("Page ");
fRun.setItalic(true);
```

## Different Headers for First Page

```csharp
var firstHeader = doc.getHeaderFooterPolicy().createHeader(XWPFHeaderFooterPolicy.FIRST);
firstHeader.createParagraph().createRun().setText("Cover Page Header");
```

## Reading Headers and Footers

```csharp
var policy = doc.getHeaderFooterPolicy();
var defaultHeader = policy.getDefaultHeader();
if (defaultHeader is not null)
{
    var text = defaultHeader.getParagraphs()[0].getText();
}

var defaultFooter = policy.getDefaultFooter();
if (defaultFooter is not null)
{
    var text = defaultFooter.getParagraphs()[0].getText();
}
```

Headers and footers are preserved on round-trip.
