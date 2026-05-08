# DotnetPoi.Ooxml

OOXML implementations for **xlsx**, **docx**, **pptx**, and macro-enabled Office 2007+ formats.

## Install

```shell
dotnet add package DotnetPoi.Ooxml
```

```xml
<PackageReference Include="DotnetPoi.Ooxml" Version="..." />
```

NuGet automatically resolves transitive dependencies (`DotnetPoi.Common`, `DotnetPoi.POIFS`).

## Support Status

`DotnetPoi.Ooxml` is the stable package for practical OOXML workflows:

| Format | Implementation | Status | Description |
|---|---|---|---|
| **xlsx / xlsm** | XSSF | ✅ Practical (~78%) | Workbook creation/read/edit, styles, layout, images, formulas-as-text, macro preservation, and Java POI interop slices. |
| **docx / docm** | XWPF | ✅ Practical (~65%) | Paragraphs/runs, tables, sections, headers/footers, images, fields, text boxes, macro preservation, and loss-resistant round-trips. |
| **pptx / pptm** | XSLF | ⚠️ Practical basic (~40%) | Slides, text, pictures, tables, slide size, layout references, macro preservation, and preservation of many advanced parts. |

### XSSF comments

xlsx cell comments support common read/create/edit/remove workflows through `XSSFComment`, cell/sheet lookup, and VML/comment part read/write. Rich comment formatting and VML shape styling are intentionally minimal and are not byte-for-byte POI-compatible.

### XWPF comments

Existing docx comments are preserved on round-trip through package part preservation. A minimal API exposes `XWPFComment` metadata/text plus document lookup, paragraph/run reference ids, metadata/text editing, and paragraph range comment creation. Rich comment body content, arbitrary range editing, and cleanup-heavy marker deletion remain limited.

### XWPF tracked changes

Tracked-change XML (`w:ins`, `w:del`, move ranges, etc.) is preserved in body/paragraph child order during round-trip. This is preservation-only: accept/reject/create/edit APIs for revisions are not modeled.

## Usage scenarios

| If you need… | Install this |
|---|---|
| xlsx / docx / pptx read/write only | `DotnetPoi.Ooxml` (this package) |
| OOXML + formula evaluator | `DotnetPoi.Ooxml` + `DotnetPoi.Formula` |
| OOXML + legacy binary (xls/doc/ppt) | `DotnetPoi.Ooxml` + `DotnetPoi.Legacy` |
| Everything (all formats + formula) | `DotnetPoi.All` (meta-package) |

## License

[Apache License 2.0](../../LICENSE) — same as upstream Apache POI.

## Disclaimer

This project is not affiliated with the Apache Software Foundation or the Apache POI project.
Apache POI is a registered trademark of the Apache Software Foundation.
