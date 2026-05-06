# Your First Word Document

This page walks through creating a `.docx` file with paragraphs, formatted runs, and an embedded image.

## Minimal Example

```csharp
using DotnetPoi.XWPF.UserModel;

using var doc = new XWPFDocument();

var paragraph = doc.createParagraph();
var run = paragraph.createRun();
run.setText("Hello from dotnet-poi!");
run.setBold(true);
run.setFontSize(14);

using var file = File.Create("output.docx");
doc.write(file);
```

## Step by Step

### 1. Create a Document

```csharp
using var doc = new XWPFDocument();
```

`XWPFDocument` represents a Word document (the `.docx` package).

### 2. Add Paragraphs and Runs

```csharp
var p1 = doc.createParagraph();
var r1 = p1.createRun();
r1.setText("Bold heading");
r1.setBold(true);
r1.setFontSize(16);

var p2 = doc.createParagraph();
var r2 = p2.createRun();
r2.setText("Italic subtitle");
r2.setItalic(true);

var p3 = doc.createParagraph();
p3.createRun().setText("Plain body text.");
```

A paragraph (`XWPFParagraph`) contains one or more runs (`XWPFRun`). Each run can have its own formatting.

### 3. Available Run Formatting

| Method | Effect |
|---|---|
| `setBold(bool)` | Bold text |
| `setItalic(bool)` | Italic text |
| `setUnderline(UnderlinePatterns)` | Underline style |
| `setStrikeThrough(bool)` | Strikethrough |
| `setFontName(string)` | Font family |
| `setFontSize(int)` | Font size in half-points |
| `setColor(string)` | Text color (hex RGB, e.g. "FF0000") |

### 4. Embed an Image

```csharp
var imageBytes = File.ReadAllBytes("photo.jpg");

var imagePara = doc.createParagraph();
var imageRun = imagePara.createRun();
var picture = imageRun.addPicture(
    imageBytes,
    XWPFPictureData.PICTURE_TYPE_JPEG,
    "photo.jpg",
    width:  914400,   // 1 inch in EMU
    height: 914400);

picture.setRotation(45.0);  // rotate 45 degrees
```

Image dimensions are in EMU (English Metric Units). 1 inch = 914400 EMU.

### 5. Save

```csharp
using var file = File.Create("output.docx");
doc.write(file);
```

## Full Runnable Example

See `examples/Phase32DocxExample/` in the repository:

[examples/Phase32DocxExample](https://github.com/modeverv/dotnet-poi/tree/main/examples/Phase32DocxExample)

Run it:

```bash
dotnet run --project examples/Phase32DocxExample
```

Generated file: `examples/output/phase3_2-docx-example.docx`

## What's Next

- [Tables in docx](../guides/docx/tables.md) — creating and reading tables
- [Headers and Footers](../guides/docx/headers-footers.md) — page headers and footers
