# DotnetPoi.Core

[![CI](https://github.com/modeverv/dotnet-poi/actions/workflows/ci.yml/badge.svg)](https://github.com/modeverv/dotnet-poi/actions/workflows/ci.yml)
[![NuGet](https://img.shields.io/nuget/v/DotnetPoi.Core)](https://www.nuget.org/packages/DotnetPoi.Core)
[![NuGet Downloads](https://img.shields.io/nuget/dt/DotnetPoi.Core)](https://www.nuget.org/packages/DotnetPoi.Core)
![License](https://img.shields.io/badge/license-Apache%202.0-blue)
![.NET Standard](https://img.shields.io/badge/.NET%20Standard-2.0-%23820171)
![Platform](https://img.shields.io/badge/platform-.NET%20%7C%20.NET%20Framework%20%7C%20Mono%20%7C%20Unity-512BD4)
![Tests](https://img.shields.io/badge/tests-244%20passing-brightgreen)

**Faithful .NET port of Apache POI — read/write xlsx, docx, pptx, xls, doc, ppt with zero dependencies.**

> ⚠️ This is an **unofficial** port and is **not affiliated with the Apache Software Foundation**. Apache POI is a registered trademark of the Apache Software Foundation.

---

## Install

```shell
dotnet add package DotnetPoi.Core
```

```xml
<PackageReference Include="DotnetPoi.Core" Version="0.5.0" />
```

> Compatible with **.NET 5+**, **.NET Core 3.1+**, **.NET Framework 4.7.2+**, **Mono 5.4+**, **Unity 2018.1+**.

---

## Documentation

Full documentation is available at the **[dotnet-poi documentation site](https://modeverv.github.io/dotnet-poi/)** 

| Section | Contents |
|---|---|
| [Getting Started](https://modeverv.github.io/dotnet-poi/getting-started/installation) | Installation, first workbook, first document, first presentation |
| [Compatibility](https://modeverv.github.io/dotnet-poi/compatibility/format-coverage) | Format coverage, limitations, package split |
| [xlsx Guides](https://modeverv.github.io/dotnet-poi/guides/xlsx/cell-types) | Cell types, styles, layout, images, formulas, data validation, conditional formatting, auto filter, pivot tables, protection, rich text, macros |
| [docx Guides](https://modeverv.github.io/dotnet-poi/guides/docx/paragraphs) | Paragraphs, tables, images, headers/footers, hyperlinks, fields, sections |
| [pptx Guides](https://modeverv.github.io/dotnet-poi/guides/pptx/slides) | Slides, images, tables, formatting |
| [Examples](https://modeverv.github.io/dotnet-poi/examples) | Runnable example projects index |

---

## Package architecture

| Package | Contents | Dependencies |
|---|---|---|
| **`DotnetPoi.Core`** (this package) | XSSF (xlsx), HSSF (xls), XWPF (docx), XSLF (pptx), HWPF (doc), HSLF (ppt), POIFS OLE2 container, common SS interfaces, `PoiXmlWriter`, OOXML Agile Encryption | **Zero** |
| `DotnetPoi.Formula` | Formula evaluator — add when you need `createFormulaEvaluator()` | References Core |

`DotnetPoi.Core` contains **everything except formula evaluation**. Formulas are preserved as text and cached values are read/written, but the evaluator engine lives in the separate `DotnetPoi.Formula` package. When `DotnetPoi.Formula` is referenced, `createFormulaEvaluator()` is automatically enabled via lazy assembly discovery at runtime.

---

## Quick start

### xlsx — write

```csharp
using DotnetPoi.XSSF.UserModel;

using var wb = new XSSFWorkbook();
var sheet = wb.createSheet("Sheet1");
var row = sheet.createRow(0);
row.createCell(0).setCellValue("Hello");
row.createCell(1).setCellValue(42);

using var fs = new FileStream("output.xlsx", FileMode.Create);
wb.write(fs);
```

### xlsx — read

```csharp
using var fs = new FileStream("input.xlsx", FileMode.Open);
using var wb = new XSSFWorkbook(fs);

var cell = wb.getSheetAt(0).getRow(0).getCell(0);
Console.WriteLine(cell.getStringCellValue());
```

### xlsx — styling

```csharp
using DotnetPoi.SS.UserModel;

var style = wb.createCellStyle();
style.setFillForegroundColor(IndexedColors.Green.getIndex());
style.setFillPattern(FillPatternType.SolidForeground);
style.setBorderTop(BorderStyle.Thin);
var cell = sheet.createRow(0).createCell(0);
cell.setCellValue("Styled");
cell.setCellStyle(style);
```

### docx — write

```csharp
using DotnetPoi.XWPF.UserModel;

using var doc = new XWPFDocument();
var run = doc.createParagraph().createRun();
run.setText("Hello from dotnet-poi");
run.setBold(true);

using var fs = new FileStream("output.docx", FileMode.Create);
doc.write(fs);
```

### pptx — write

```csharp
using DotnetPoi.XSLF.UserModel;

using var prs = new XMLSlideShow();
var slide = prs.createSlide();
var picIdx = prs.addPicture(File.ReadAllBytes("photo.jpeg"), XSLFPictureData.PICTURE_TYPE_JPEG);
var shape = prs.createPicture(slide, picIdx);
shape.setAnchor(0, 0, XMLSlideShow.DefaultSlideCx, XMLSlideShow.DefaultSlideCy);
shape.setRotation(45.0);

using var fs = new FileStream("output.pptx", FileMode.Create);
prs.write(fs);
```

### Encrypted xlsx

```csharp
using var wb = new XSSFWorkbook();
wb.createSheet("Sheet1").createRow(0).createCell(0).setCellValue("Secret");
using var fs = new FileStream("secret.xlsx", FileMode.Create);
wb.writeEncrypted(fs, "password123");
```

---

## Format coverage

Legend: ✅ complete / ⚠️ partial / 🔵 preserved as unknown parts, not editable / ❌ not implemented.

### xlsx / XSSF (~78%)

| Category | Feature | Status |
|---|---|---|
| **Cell values** | string, numeric, date, boolean, error | ✅ |
| **Formulas** | formula text write/read + cached value read | ✅ |
| **Formulas** | programmatic formula evaluation | ❌ — see DotnetPoi.Formula |
| **Styles** | fonts (name, size, bold, italic, color, underline, strikethrough) | ✅ |
| **Styles** | fill (pattern + foreground colour) | ✅ |
| **Styles** | borders (4 sides, each with style+colour) | ✅ |
| **Styles** | number/date format codes | ✅ |
| **Styles** | alignment (horizontal, vertical, wrap, indent, rotation) | ✅ |
| **Layout** | merged cells, column width, row height | ✅ |
| **Layout** | hidden rows/columns | ✅ |
| **Layout** | freeze panes (pane splits) | ✅ |
| **Layout** | active sheet index, active cell API | ✅ |
| **Layout** | print settings (margins, page size, orientation, header/footer) | ✅ |
| **Drawings** | images (multiple, anchor, rotation) | ✅ |
| **Drawings** | hyperlinks | ✅ |
| **Drawings** | charts | 🔵 preserved, not creatable |
| **Drawings** | comments | 🔵 preserved, not creatable |
| **Drawings** | auto-shapes | 🔵 | Unknown `xdr:twoCellAnchor` children preserved via raw XML capture/re-emission |
| **Data** | data validation (input rules) | ✅ |
| **Data** | conditional formatting | ✅ |
| **Data** | auto filter | ✅ |
| **Data** | pivot tables | ⚠️ creation only, editing not modelled |
| **Strings** | shared strings (plain) | ✅ |
| **Strings** | rich text (per-character formatting) | ✅ |
| **Other** | workbook/sheet protection | ✅ |
| **Other** | macro-enabled (xlsm) | ✅ VBA bytes preserved |
| **Other** | sparklines | ❌ |
| **Other** | external data connections | 🔵 | `xl/connections.xml` / `xl/externalLinks/*` round-trip via `_preservedEntries` |

### docx / XWPF (~65%)

| Category | Feature | Status |
|---|---|---|
| **Paragraphs/runs** | text, font name/size/color, bold, italic, underline, strikethrough | ✅ |
| **Paragraphs** | alignment, indents, spacing, bullet/numbered lists | ✅ |
| **Tables** | create/read tables, rows, cells | ✅ |
| **Tables** | cell merging, borders | 🔵 | Round-trip preserved via raw XML capture/re-emission; API-level creation not modeled |
| **Sections** | page setup, headers, footers | ✅ |
| **Sections** | columns | ✅ | `setColumns()` API, round-trip verified |
| **Links** | external hyperlinks | ✅ |
| **Images** | inline images + rotation | ✅ |
| **Images** | text boxes (`w:txbxContent`) | ❌ |
| **Review** | comments, footnotes, endnotes | 🔵 | Existing parts round-trip via `_preservedEntries` |
| **Fields** | TOC, page numbers, mail merge fields | ✅ |
| **Styles** | paragraph style reference (pStyle) | ✅ | `setStyle()`/`getStyleID()` API, round-trip verified. Character/table styles ❌. `word/styles.xml` 🔵 preserved + default styles auto-generated. |
| **Other** | macro-enabled (docm) | ✅ |
| **Other** | content controls (SDT) | 🔵 | Block-level and inline SDT preserved via raw XML capture/re-emission |
| 〃 | tracked changes (ins/del/move) | ❌ | |
| 〃 | OLE embeddings | 🔵 | `word/embeddings/*` round-trip via `_preservedEntries` |

### pptx / XSLF (~40%)

| Category | Feature | Status |
|---|---|---|
| **Slides** | create/read, slide size | ✅ |
| **Slides** | notes slides | 🔵 | `ppt/notesSlides/*` round-trip via `_preservedEntries` |
| **Text** | text boxes, multiple paragraphs, run formatting | ✅ |
| **Shapes** | pictures (anchor, size, rotation) | ✅ |
| **Shapes** | tables (`p:graphicFrame` / `a:tbl`) | ✅ |
| **Shapes** | group shapes, connectors, lines | 🔵 | Unknown `p:spTree` children preserved via raw XML capture/re-emission |
| **Shapes** | SmartArt, charts | 🔵 preserved, not modelled |
| **Media** | video/audio embedding | 🔵 | Non-image `ppt/media/*` round-trip via `_preservedEntries` |
| **Animation** | animations, transitions | 🔵 preserved, not modelled |
| **Theme** | layouts, masters, themes | 🔵 preserved, not editable |
| **Other** | macro-enabled (pptm) | ✅ |

### xls / HSSF (~10%)

| Category | Feature | Status |
|---|---|---|
| **Basic** | read/write cell values | ⚠️ in progress |
| **Everything else** | BIFF records, styles, formulas, images, filters, pivots | ❌ |

### Legacy binary

| Format | Status |
|---|---|
| doc (HWPF) | ~5% read-only MVP |
| ppt (HSLF) | ~5% read-only MVP |

---

## Practical gaps (priority order)

| # | Gap | Why it matters |
|---|---|---|
| 1 | Formula evaluation | Template fill → save → open in Excel works, but programmatic access to freshly calculated values needs DotnetPoi.Formula |
| 2 | Chart creation | Reports and presentations commonly need charts from data |
| 3 | docx styles | Word documents rely on named styles (Normal, Heading1…) for formatting |
| 4 | HSSF `.xls` depth | Legacy format support is still minimal |
| 5 | docx text boxes | Word text inside `w:txbxContent` elements is not read |

---

## Test coverage

| Project | Tests | Notes |
|---|---|---|
| Core.Tests | 244 | xlsx round-trip coverage, active sheet, auto filter, protection, fields, docx table/rPr preservation |
| Formula.Tests | 10 | Minimal — evaluator is early stage |
| Interop.Tests (C#) | 55 | Bidirectional Java/.NET fixtures + preservation tests |
| **Total (C#)** | **309** | |
| Java POI side (Maven) | 44 | Fixture generation / readback |

---

## Philosophy

- 🔱 Maximum fidelity to upstream Apache POI — we follow, not reinvent
- 🤖 Ported class-by-class with LLM assistance
- 💸 Free forever. No EULA. No maintenance fee. No exceptions.
- 📖 Apache POI is the source of truth — included as a git submodule

---

## License

[Apache License 2.0](../../LICENSE) — same as upstream Apache POI.

## Disclaimer

This project is not affiliated with the Apache Software Foundation or the Apache POI project.
Apache POI is a registered trademark of the Apache Software Foundation.
