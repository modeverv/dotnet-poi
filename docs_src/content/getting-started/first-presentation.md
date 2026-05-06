# Your First PowerPoint Presentation

This page walks through creating a `.pptx` file with slides, text boxes, images, and tables.

## Minimal Example

```csharp
using DotnetPoi.XSLF.UserModel;

using var prs = new XMLSlideShow();

var slide = prs.createSlide();

// Add a text box
var textBox = slide.createTextBox();
textBox.setText("Hello from dotnet-poi!");
textBox.setAnchor(100, 100, 400, 200);

using var file = File.Create("output.pptx");
prs.write(file);
```

## Step by Step

### 1. Create a Presentation

```csharp
using var prs = new XMLSlideShow();
```

`XMLSlideShow` represents a PowerPoint presentation (the `.pptx` package).

### 2. Add Slides

```csharp
var slide1 = prs.createSlide();
var slide2 = prs.createSlide();
```

Each call to `createSlide()` appends a new blank slide.

### 3. Add a Text Box

```csharp
var textBox = slide.createTextBox();
textBox.setText("Hello, World!");
```

Set position and size with `setAnchor` (x, y, width, height in EMU):

```csharp
textBox.setAnchor(100, 100, 400, 200);
```

Format individual text runs:

```csharp
var tp = textBox.addNewTextParagraph();
var tr = tp.addNewTextRun();
tr.setText("Bold title");
tr.setBold(true);
tr.setFontSize(24);
tr.setFontColor("FF0000");
```

### 4. Add an Image

```csharp
var imageBytes = File.ReadAllBytes("photo.jpg");
var picIdx = prs.addPicture(imageBytes, XSLFPictureData.PICTURE_TYPE_JPEG);

var slide = prs.createSlide();
var shape = prs.createPicture(slide, picIdx);
shape.setAnchor(100, 100, 2000, 1500);
shape.setRotation(45.0);
```

`addPicture` returns a picture index. Use it with `createPicture` to place the image on a slide.

### 5. Add a Table

```csharp
var table = slide.createTable(3, 2);
table.setAnchor(100, 100, 600, 200);

table.getCell(0, 0).setText("Name");
table.getCell(0, 1).setText("Value");
table.getCell(1, 0).setText("Item 1");
table.getCell(1, 1).setText("42");
```

### 6. Save

```csharp
using var file = File.Create("output.pptx");
prs.write(file);
```

## Full Runnable Example

See `examples/Phase33PptxExample/` in the repository:

[examples/Phase33PptxExample](https://github.com/modeverv/dotnet-poi/tree/main/examples/Phase33PptxExample)

Run it:

```bash
dotnet run --project examples/Phase33PptxExample
```

Generated file: `examples/output/phase3_3-pptx-example.pptx`

## What's Next

- [Images in pptx](../guides/pptx/images.md) — placement, rotation, scaling
- [Tables in pptx](../guides/pptx/tables.md) — creating styled tables
