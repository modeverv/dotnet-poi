# Headers and Footers

Add headers and footers to docx documents.

## Adding a Header

```csharp
using DotnetPoi.XWPF.UserModel;

using var doc = new XWPFDocument();
doc.setHeaderText("Document Header");
```

## Adding a Footer

```csharp
doc.setFooterText("Page footer");
```

## Different Headers for First Page

```csharp
doc.setFirstHeaderText("Cover Page Header");
doc.setEvenHeaderText("Even Page Header");
doc.setFirstFooterText("Cover Page Footer");
doc.setEvenFooterText("Even Page Footer");
```

## Reading Headers and Footers

```csharp
var defaultHeader = doc.getHeaderText();
var firstHeader = doc.getFirstHeaderText();
var evenFooter = doc.getEvenFooterText();
```

Headers and footers are preserved on round-trip.

## Full Runnable Example

See `examples/UsageSamples/Program.cs` (`CreateDocument`):

[examples/UsageSamples](https://github.com/modeverv/dotnet-poi/tree/master/examples/UsageSamples)
