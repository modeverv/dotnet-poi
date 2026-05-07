# DotnetPoi.Ooxml

OOXML implementations for **xlsx**, **docx**, **pptx**, and macro-enabled Office 2007+ formats.

`DotnetPoi.Ooxml` 1.0 is stable for the documented common OOXML workflows. It is not a full Apache POI surface clone: advanced features such as chart creation, comment editing, SmartArt, animations, tracked changes, and some deep style models remain limited or preservation-only.

## Install

```shell
dotnet add package DotnetPoi.Ooxml --version 1.0.0
```

```xml
<PackageReference Include="DotnetPoi.Ooxml" Version="1.0.0" />
```

NuGet automatically resolves transitive dependencies (`DotnetPoi.Common`, `DotnetPoi.POIFS`).

## Usage scenarios

| If you need… | Install this |
|---|---|
| xlsx / docx / pptx read/write only | `DotnetPoi.Ooxml` (this package) |
| OOXML + formula evaluator | `DotnetPoi.Ooxml` + `DotnetPoi.Formula` |
| OOXML + legacy binary (xls/doc/ppt) | `DotnetPoi.Ooxml` + `DotnetPoi.Legacy` |
| Everything (all formats + formula) | `DotnetPoi.All` (meta-package) |

## 1.0 support scope

| Format | 1.0 status |
|---|---|
| **xlsx / xlsm** | Strongest surface: workbook creation, read/edit/write, styles, layout, images, formula text/cached values, protection, macro preservation, and Java POI interop coverage. |
| **docx / docm** | Practical document generation and light editing: paragraphs, runs, tables, sections, headers/footers, hyperlinks, images, fields, macro and unknown-part preservation. |
| **pptx / pptm** | Practical simple presentation generation: slides, text boxes, formatted runs, pictures, rotation, tables, slide size, macro and unknown-part preservation. Advanced PowerPoint objects are mostly preservation-only. |

Formula evaluation is intentionally separate. Add `DotnetPoi.Formula` only when you need the supported evaluator subset; formula text read/write and cached values work without it.

## License

[Apache License 2.0](../../LICENSE) — same as upstream Apache POI.

## Disclaimer

This project is not affiliated with the Apache Software Foundation or the Apache POI project.
Apache POI is a registered trademark of the Apache Software Foundation.
