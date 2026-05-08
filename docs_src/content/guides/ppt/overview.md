# ppt (HSLF) Overview

HSLF is the legacy PowerPoint 97-2003 `.ppt` implementation. Coverage is ~12%.

## Status

HSLF is currently a preservation-first reader. It can open OLE2 `.ppt` files, inventory streams, read the `PowerPoint Document` record tree, recover slide order, extract text from `TextCharsAtom` and `TextBytesAtom` records, and no-op write the compound file back out. This is useful for archive discovery and cautious round-trip preservation, but it is not a slide authoring or shape editing engine.

The next planned work is to deepen Java POI interop assertions and add a usermodel only where preservation is already stable.

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
- OLE2 stream inventory
- `PowerPoint Document` record tree scan with raw record preservation
- Slide count and slide order reconstruction
- Text extraction from `TextCharsAtom` and `TextBytesAtom`
- No-op write / round-trip preservation

## Limitations

- Java POI direction-B assertion is still pending
- No shapes, images, comments, notes, animations, transitions, or master/layout editing model
- No public editing API

For new presentations, prefer `pptx` / XSLF.
