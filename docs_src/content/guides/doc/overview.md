# doc (HWPF) Overview

HWPF provides early support for the legacy Word 97-2003 `.doc` format. Coverage is ~20%.

## Status

HWPF is useful today for body-text extraction, simple indexing/migration workflows, and limited main-body edits. It can open OLE2 `.doc` files, parse the File Information Block and selected table stream, extract main document text from the CLX/piece table, expose a minimal Range/Paragraph/CharacterRun model, and preserve unedited OLE streams/storages during no-op or limited body edits.

It is not a complete Word binary editing engine. Tables, images, header/footer stories, footnotes, comments, fields, and tracked changes are not modeled through public APIs.

## Basic Text Extraction

```csharp
using DotnetPoi.HWPF.UserModel;

using var stream = File.OpenRead("input.doc");
using var doc = new HWPFDocument(stream);

Console.WriteLine(doc.getText());
```

## Limited Body Editing

```csharp
using DotnetPoi.HWPF.UserModel;

using var stream = File.OpenRead("input.doc");
using var doc = new HWPFDocument(stream);

doc.appendParagraph("Added by dotnet-poi");
doc.replaceText("{{name}}", "Example Corp");

using var output = File.Create("output.doc");
doc.write(output);
```

## Supported Today

- OLE2 `.doc` open and stream inventory
- FIB / `0Table` / `1Table` parsing
- Main body text extraction from CLX/piece table
- Compressed and Unicode text pieces
- Minimal Range / Paragraph / CharacterRun API
- Selected CHPX character formatting: font name, size, bold, italic, underline, strike
- Minimal paragraph property reading
- No-op write preservation
- Limited append paragraph and simple text replacement
- OLE stream/storage and embedded OLE preservation
- Java POI Direction B smoke for dotnet-poi no-op saved `.doc`

## Limitations

- No table model or table editing API
- No image extraction or editing API
- No header/footer story extraction API
- No footnote/endnote/comment model
- No field/bookmark model
- No full Word style inheritance or complete PAPX/CHPX expansion
- Limited body edits rebuild the main body as a single Unicode piece; use cautiously for complex documents

For new documents, prefer `docx` / XWPF.
