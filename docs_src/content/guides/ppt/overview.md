# ppt (HSLF) Overview

HSLF is the legacy PowerPoint 97-2003 `.ppt` implementation. Coverage is ~5%.

## Status

HSLF is currently a minimal reader. It can open OLE2 `.ppt` files, read the `PowerPoint Document` stream, and extract early text content from `TextCharsAtom` and `TextBytesAtom` records. This can be useful for very basic text discovery experiments, but it is not yet ready for reliable migration, preservation, or editing workflows.

The next planned work is to add fixture coverage, OLE2 stream inventory, record tree preservation, no-op write round-trip, and Java POI interop.

## Basic Text Probe

```csharp
using DotnetPoi.HSLF.UserModel;

using var stream = File.OpenRead("input.ppt");
using var ppt = new HSLFSlideShow(stream);

foreach (var slide in ppt.getSlides())
{
    foreach (var text in slide.getTextParagraphs())
        Console.WriteLine(text);
}
```

## Supported Today

- OLE2 `.ppt` open
- `PowerPoint Document` stream scan
- Minimal slide container detection
- Text extraction from `TextCharsAtom` and `TextBytesAtom`

## Limitations

- No no-op write preservation yet
- No Java POI interop track yet
- Slide order is not yet rebuilt from `Current User` / `UserEditAtom` / persist pointers
- No shapes, images, comments, notes, animations, transitions, OLE, or master/layout model
- No public editing API

For new presentations, prefer `pptx` / XSLF.
