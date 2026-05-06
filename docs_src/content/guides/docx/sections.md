# Page Setup (Sections)

Control page size, margins, and orientation in docx documents.

## Page Size

```csharp
using DotnetPoi.XWPF.UserModel;

using var doc = new XWPFDocument();

var section = doc.getDocument().getBody().addNewSectPr();
var pgSz = section.addNewPgSz();
pgSz.w = 11906;   // A4 width in twips (1/1440 inch)
pgSz.h = 16838;   // A4 height
```

Common paper sizes in twips:

| Paper | Width | Height |
|---|---|---|
| A4 | 11906 | 16838 |
| Letter | 12240 | 15840 |
| A3 | 16838 | 23811 |

## Orientation

```csharp
pgSz.w = 16838;   // swap width/height for landscape
pgSz.h = 11906;
```

## Margins

```csharp
var pgMar = section.addNewPgMar();
pgMar.left   = 1440;   // 1 inch in twips
pgMar.right  = 1440;
pgMar.top    = 1440;
pgMar.bottom = 1440;
pgMar.header = 720;    // 0.5 inch
pgMar.footer = 720;
```

1 inch = 1440 twips. 1 cm ≈ 567 twips.

## Reading Page Settings

```csharp
var sectPr = doc.getDocument().getBody().getSectPr();
var size = sectPr.getPgSz();
Console.WriteLine($"Width: {size.w}, Height: {size.h}");
```

Page settings are preserved on round-trip.
