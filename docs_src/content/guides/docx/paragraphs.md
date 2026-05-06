# Paragraphs and Runs

Create and format paragraphs in a docx document.

## Basic Paragraph

```csharp
using DotnetPoi.XWPF.UserModel;

using var doc = new XWPFDocument();

var p = doc.createParagraph();
var r = p.createRun();
r.setText("Hello from dotnet-poi!");
```

## Run Formatting

```csharp
var r = paragraph.createRun();
r.setText("Formatted text");
r.setBold(true);
r.setItalic(true);
r.setFontName("Arial");
r.setFontSize(16);
r.setColor("FF0000");             // hex RGB
r.setUnderline(UnderlinePatterns.Single);
r.setStrikeThrough(true);
```

## Paragraph Alignment

```csharp
paragraph.setAlignment(ParagraphAlignment.CENTER);
// LEFT, CENTER, RIGHT, BOTH (justified)
```

## Indentation and Spacing

```csharp
paragraph.setIndentationLeft(500);    // left indent in twips (1/1440 inch)
paragraph.setIndentationRight(300);
paragraph.setIndentationFirstLine(200);  // first-line indent
paragraph.setSpacingBefore(200);         // space before in twips
paragraph.setSpacingAfter(100);
paragraph.setSpacingBetween(1.5);        // line spacing (1.0, 1.5, 2.0)
```

## Bullet and Numbered Lists

```csharp
var bullet = doc.createParagraph();
bullet.setBullet(true);   // bullet point
bullet.createRun().setText("First item");

var numbered = doc.createParagraph();
numbered.setNumbering(1);  // numbered (1, 2, 3...)
numbered.createRun().setText("Step one");
```

## Full Runnable Example

See `examples/Phase32DocxExample/` and `examples/UsageSamples/Program.cs` (`CreateDocument`):

[examples/Phase32DocxExample](https://github.com/modeverv/dotnet-poi/tree/main/examples/Phase32DocxExample)

[examples/UsageSamples](https://github.com/modeverv/dotnet-poi/tree/main/examples/UsageSamples)
