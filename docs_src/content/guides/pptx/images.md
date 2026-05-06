# Images

Embed images on pptx slides with position, size, rotation, and flipping.

## Adding a Picture

```csharp
using DotnetPoi.XSLF.UserModel;

using var ppt = new XMLSlideShow();
var imageBytes = File.ReadAllBytes("photo.jpg");

// Add picture to the presentation (returns the index)
var imageIndex = ppt.addPicture(imageBytes, XSLFPictureData.PICTURE_TYPE_JPEG);

// Create picture on a slide
var slide = ppt.createSlide();
var picture = ppt.createPicture(slide, imageIndex);
picture.setAnchor(914400, 685800, 4572000, 3429000);
// setAnchor(x, y, width, height) in EMU
```

## Rotation

```csharp
picture.setRotation(8.0);  // rotate 8 degrees (decimal degrees)
```

## Flipping

```csharp
picture.FlipHorizontal = true;
picture.FlipVertical = true;
```

## Supported Image Types

| Constant | Format |
|---|---|
| `PICTURE_TYPE_JPEG` | JPEG |
| `PICTURE_TYPE_PNG` | PNG |
| `PICTURE_TYPE_GIF` | GIF |
| `PICTURE_TYPE_BMP` | BMP |

## Picture with Text on the Same Slide

```csharp
// Add picture
var picture = ppt.createPicture(slide, imageIndex);
picture.setAnchor(914400, 685800, 4572000, 3429000);

// Add text box alongside
var textBox = slide.createTextBox();
textBox.setAnchor(5715000, 914400, 2743200, 1371600);
textBox.addParagraph().addRun("Image caption");
```

## Full Runnable Example

See `examples/Phase33PptxExample/` and `examples/UsageSamples/Program.cs` (`CreatePresentation`):

[examples/Phase33PptxExample](https://github.com/modeverv/dotnet-poi/tree/main/examples/Phase33PptxExample)

[examples/UsageSamples](https://github.com/modeverv/dotnet-poi/tree/main/examples/UsageSamples)
