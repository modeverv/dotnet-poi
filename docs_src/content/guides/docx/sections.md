# Page Setup and Sections

Control page size, margins, orientation, and section columns in docx documents.

## Page Size

```csharp
using DotnetPoi.XWPF.UserModel;

using var doc = new XWPFDocument();
doc.setPageSize(11906, 16838);   // A4 in twips (1/1440 inch)
```

Common paper sizes in twips:

| Paper | Width | Height |
|---|---|---|
| A4 | 11906 | 16838 |
| Letter | 12240 | 15840 |
| A3 | 16838 | 23811 |

## Orientation

```csharp
doc.setPageSize(16838, 11906);
doc.setLandscape(true);
```

## Margins

```csharp
doc.setMargins(
    top: 1440,
    right: 1440,
    bottom: 1440,
    left: 1440);
```

1 inch = 1440 twips. 1 cm ≈ 567 twips.

## Columns

```csharp
doc.setColumns(count: 2, spacingTwips: 720);

var count = doc.getColumnCount();
var spacing = doc.getColumnSpacing();
```

## Reading Page Settings

```csharp
Console.WriteLine($"{doc.getPageWidth()} x {doc.getPageHeight()}");
Console.WriteLine(doc.isLandscape());
Console.WriteLine(doc.getMarginLeft());
```

Page settings, final-section columns, and raw paragraph-level `sectPr` section breaks are preserved on round-trip.

## Full Runnable Example

See `examples/UsageSamples/Program.cs` (`CreateDocument`):

[examples/UsageSamples](https://github.com/modeverv/dotnet-poi/tree/main/examples/UsageSamples)
