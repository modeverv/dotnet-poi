# DotnetPoi.Legacy

Legacy Office 97–2003 binary format implementations for **xls** (HSSF), **doc** (HWPF), and **ppt** (HSLF).

## Install

```shell
dotnet add package DotnetPoi.Legacy
```

```xml
<PackageReference Include="DotnetPoi.Legacy" Version="..." />
```

NuGet automatically resolves transitive dependencies (`DotnetPoi.Common`, `DotnetPoi.POIFS`).

## Support Status

| Format | Implementation | Status | Description |
|---|---|---|---|
| **xls** | HSSF | ⚠️ Partial (~35%) | Basic workbook read/write, styles, layout, and OLE/BIFF preservation. |
| **doc** | HWPF | ⚠️ Partial (~20%) | Main body text extraction and limited body editing with OLE preservation. |
| **ppt** | HSLF | ❌ Experimental (~5%) | Early text extraction reader; no write support yet. |

### HSSF (Excel 97-2003)

- **Supported:** String, numeric, boolean, error cells; multiple sheets; basic fonts/styles; column/row sizing; merged regions; freeze panes; reading formula text/values; preservation of macros (VBA), OLE streams, and unknown BIFF records.
- **Not yet modeled:** Image/shape/chart creation; comment/hyperlink editing; filters; pivots; new formula token writing.

### HWPF (Word 97-2003)

- **Supported:** OLE2 open; FIB parsing; main body text extraction (CLX/piece table); header/footer and table text extraction; basic Range/Paragraph/Run model; bold/italic/size/font; no-op write; append paragraph; simple text replacement; OLE preservation.
- **Not yet modeled:** Images; footnotes; comments; fields; bookmarks.

### HSLF (PowerPoint 97-2003)

- **Supported:** OLE2 open; `PowerPoint Document` stream scan; basic text extraction.
- **Not yet modeled:** Writing; slide order reconstruction; shapes; images; master slides.

## Usage scenarios

| If you need… | Install this |
|---|---|
| xls / doc / ppt read/write only | `DotnetPoi.Legacy` (this package) |
| Legacy binary + formula evaluator | `DotnetPoi.Legacy` + `DotnetPoi.Formula` |
| Legacy binary + OOXML (xlsx/docx/pptx) | `DotnetPoi.Legacy` + `DotnetPoi.Ooxml` |
| Everything (all formats + formula) | `DotnetPoi.All` (meta-package) |

## License

[Apache License 2.0](../../LICENSE) — same as upstream Apache POI.

## Disclaimer

This project is not affiliated with the Apache Software Foundation or the Apache POI project.
Apache POI is a registered trademark of the Apache Software Foundation.
