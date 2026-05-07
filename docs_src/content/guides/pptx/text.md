# Text Boxes

Add and format text on slides.

## Creating a Text Box

```csharp
using DotnetPoi.XSLF.UserModel;

using var ppt = new XMLSlideShow();
var slide = ppt.createSlide();
var textBox = slide.createTextBox();
textBox.setAnchor(685800, 457200, 7315200, 914400);
// setAnchor(x, y, width, height) — all values in EMU
```

## Adding Text

```csharp
var paragraph = textBox.addParagraph();
var run = paragraph.addRun("Hello from dotnet-poi!");
```

## Multiple Paragraphs

```csharp
var titleBox = slide.createTextBox();
titleBox.setAnchor(685800, 457200, 7315200, 914400);
var title = titleBox.addParagraph().addRun("Title");
title.Bold = true;
title.FontSize = 28;

var subtitleBox = slide.createTextBox();
subtitleBox.setAnchor(685800, 1371600, 7315200, 914400);
subtitleBox.addParagraph().addRun("This is a subtitle.");
```

## Run Formatting

```csharp
var run = paragraph.addRun("Formatted text");
run.Bold = true;
run.Italic = true;
run.FontSize = 24;
run.FontName = "Arial";
run.FontColor = "FF0000";  // hex RGB
run.Underline = true;
```

## Reading Slide Text

```csharp
var shapes = slide.getAutoShapes();  // returns text boxes
var text = shapes[0].Paragraphs[0].getPlainText();
```

## Full Runnable Example

See `examples/Phase33PptxExample/` and `examples/UsageSamples/Program.cs` (`CreatePresentation`):

[examples/Phase33PptxExample](https://github.com/modeverv/dotnet-poi/tree/master/examples/Phase33PptxExample)

[examples/UsageSamples](https://github.com/modeverv/dotnet-poi/tree/master/examples/UsageSamples)
