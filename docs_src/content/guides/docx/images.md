# Images

Embed images in docx documents with position, size, and rotation.

## Embedding an Image

```csharp
using DotnetPoi.XWPF.UserModel;

using var doc = new XWPFDocument();
var imageBytes = File.ReadAllBytes("photo.jpg");

var imagePara = doc.createParagraph();
var imageRun = imagePara.createRun();
var picture = imageRun.addPicture(
    imageBytes,
    XWPFPictureData.PICTURE_TYPE_JPEG,
    "photo.jpg",
    width:  914400,    // 1 inch in EMU
    height: 914400);

picture.setRotation(45.0);  // rotate 45 degrees
```

## Image Dimensions

Dimensions are in EMU (English Metric Units):

| Unit | EMU |
|---|---|
| 1 inch | 914400 |
| 1 cm | 360000 |
| 1 point | 12700 |

## Supported Image Types

| Constant | Format |
|---|---|
| `PICTURE_TYPE_JPEG` | JPEG |
| `PICTURE_TYPE_PNG` | PNG |
| `PICTURE_TYPE_GIF` | GIF |
| `PICTURE_TYPE_BMP` | BMP |

## Reading Images Back

```csharp
var allPictures = doc.getAllPictures();
var pic = allPictures[0];
var picType = pic.getPictureType();  // PICTURE_TYPE_JPEG

// Check image count per run
var runs = paragraphs[0].getRuns();
var embedded = runs[0].getEmbeddedPictures();
var rotation = embedded[0].getRotation();
```

## Full Runnable Example

See `examples/Phase32DocxExample/` and `examples/UsageSamples/Program.cs` (`CreateDocument`):

[examples/Phase32DocxExample](https://github.com/modeverv/dotnet-poi/tree/main/examples/Phase32DocxExample)

[examples/UsageSamples](https://github.com/modeverv/dotnet-poi/tree/main/examples/UsageSamples)
