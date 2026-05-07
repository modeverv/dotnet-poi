# Images

Embed images into xlsx sheets with position, size, and rotation.

## Embedding an Image

```csharp
using DotnetPoi.XSSF.UserModel;

using var wb = new XSSFWorkbook();
var sheet = wb.createSheet("Images");

// Read image bytes
var imageBytes = File.ReadAllBytes("photo.jpg");
var pictureIndex = wb.addPicture(imageBytes, XSSFWorkbook.PICTURE_TYPE_JPEG);

// Create drawing
var drawing = sheet.createDrawingPatriarch();

// Create anchor (client anchor: rows 0-3, columns 0-3, with dx/dy offsets)
var anchor = drawing.createAnchor(0, 0, 0, 0, 0, 0, 3, 3);

// Create picture shape
var picture = drawing.createPicture(anchor, pictureIndex);
picture.setRotation(45.0);  // rotate 45 degrees
```

## Anchor Types

Use `createAnchor(dx1, dy1, dx2, dy2, col1, row1, col2, row2)`:

| Parameter | Meaning |
|---|---|
| `dx1, dy1` | Offset from top-left of cell (col1, row1) in EMU |
| `col1, row1` | Top-left cell coordinate |
| `dx2, dy2` | Offset from top-left of cell (col2, row2) in EMU |
| `col2, row2` | Bottom-right cell coordinate |

1 EMU = 1/914400 inch. 1 inch = 914400 EMU.

## Multiple Images

Add multiple pictures to the same sheet by calling `createPicture` multiple times:

```csharp
var pic1 = drawing.createPicture(anchor1, picIdx1);
var pic2 = drawing.createPicture(anchor2, picIdx2);
```

Each picture needs its own anchor and unique picture index from `addPicture`.

## Full Runnable Example

See `examples/Phase25ImagesExample/`:

[examples/Phase25ImagesExample](https://github.com/modeverv/dotnet-poi/tree/master/examples/Phase25ImagesExample)
