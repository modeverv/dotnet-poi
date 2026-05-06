# dotnet-poi

An **unofficial**, faithful port of [Apache POI](https://poi.apache.org/) for .NET.

[![CI](https://github.com/modeverv/dotnet-poi/actions/workflows/ci.yml/badge.svg)](https://github.com/modeverv/dotnet-poi/actions/workflows/ci.yml)
[![Examples](https://github.com/modeverv/dotnet-poi/actions/workflows/examples.yml/badge.svg)](https://github.com/modeverv/dotnet-poi/actions/workflows/examples.yml)
[![XML Parity](https://github.com/modeverv/dotnet-poi/actions/workflows/xml-parity-fixtures.yml/badge.svg)](https://github.com/modeverv/dotnet-poi/actions/workflows/xml-parity-fixtures.yml)
[![NuGet Core](https://img.shields.io/nuget/v/DotnetPoi.Core)](https://www.nuget.org/packages/DotnetPoi.Core)
[![NuGet Core Downloads](https://img.shields.io/nuget/dt/DotnetPoi.Core)](https://www.nuget.org/packages/DotnetPoi.Core)
[![NuGet Formula](https://img.shields.io/nuget/v/DotnetPoi.Formula)](https://www.nuget.org/packages/DotnetPoi.Formula)
[![NuGet Formula Downloads](https://img.shields.io/nuget/dt/DotnetPoi.Formula)](https://www.nuget.org/packages/DotnetPoi.Formula)
![License](https://img.shields.io/badge/license-Apache%202.0-blue)
![Status](https://img.shields.io/badge/status-beta-orange)
![Tests](https://img.shields.io/badge/tests-309%20passing-brightgreen)

## NuGet Package Strategy

dotnet-poi ships as **two separate NuGet packages** with a clear separation of concerns:

| Package | Contents | When to use |
|---|---|---|
| **DotnetPoi.Core** | All format implementations (XSSF/xlsx, HSSF/xls, XWPF/docx, XSLF/pptx, HWPF/doc, HSLF/ppt, POIFS) + common interfaces + XML writer | Always — required for any read/write operation |
| **DotnetPoi.Formula** | Formula evaluator (`IFormulaEvaluator`, `FormulaEvaluator`, `CellValue`) | Only when you need spreadsheet formula evaluation |

**Design principle:** `Core` has zero knowledge of `Formula`. Adding `DotnetPoi.Formula` to your project automatically enables `createFormulaEvaluator()` via lazy assembly discovery at runtime. Without it, the call throws a clear `NotSupportedException`.

```xml
<!-- Simple read/write only -->
<PackageReference Include="DotnetPoi.Core" Version="..." />

<!-- Add formula evaluation -->
<PackageReference Include="DotnetPoi.Core" Version="..." />
<PackageReference Include="DotnetPoi.Formula" Version="..." />
```

**Why this split:**
- **Core can be stable early** — all spreadsheet read/write/format logic is self-contained. No dependency on the formula engine.
- **Formula can mature slowly** — evaluation is complex and can be iterated independently without affecting Core.
- **Smaller dependency for simple use cases** — users who only need xlsx read/write don't pull in the entire formula engine.
- **Security** — applications handling untrusted documents can omit the formula evaluator entirely, reducing the attack surface.

## Philosophy

- 🔱 Maximum fidelity to upstream Apache POI — we follow, not reinvent
- 🤖 Ported class-by-class with LLM assistance, with tests written alongside
- 💸 Free forever. No EULA. No maintenance fee. No exceptions.
- 📖 Apache POI is the source of truth — included as a git submodule
- ⚠️ Not affiliated with the Apache Software Foundation

---

## Status

Current status: **beta** — `DotnetPoi.Core` v0.1.0 and `DotnetPoi.Formula` v0.1.0 are published on [NuGet.org](https://www.nuget.org/packages/DotnetPoi.Core).

The strongest format today is **xlsx / XSSF**. It has broad support for common workbook creation, reading, editing, styling, layout, images, formulas-as-text, macro preservation, and Java POI interop. Some advanced workbook features are still preservation-only rather than editable API models, but ordinary `.xlsx` workflows are well beyond the original bootstrap stage.

Legend: ✅ complete / ⚠️ partial / 🔵 preserved as unknown parts, but not modeled for creation or editing / ❌ not implemented / — not applicable.

### Format Coverage

#### xlsx / XSSF (~78%)

| Category | Feature | Status | Notes |
|---|---|---|---|
| Cell values | string, numeric, date, boolean, error | ✅ | |
| Formulas | formula text write/read + cached value read | ✅ | Excel recalculation-on-open workflow is supported. |
| Formulas | full formula evaluation | ❌ deferred | Programmatic access to freshly calculated results is not a current goal. |
| Styles | fonts, fills, borders, number/date formats, alignment | ✅ | Round-trip covered. |
| Layout | merged cells, column width, row height, hidden rows/columns, freeze panes, print settings | ✅ | |
| Drawings | images, anchors, rotation, hyperlinks | ✅ | |
| Drawings | charts, comments | 🔵 | Existing package parts are preserved during round-trip; new creation/editing is not modeled. |
| Drawings | auto-shapes | 🔵 | Unknown `xdr:twoCellAnchor` children (auto-shapes, connectors, group shapes) are preserved verbatim via raw XML capture/re-emission in drawing.xml. |
| Data | data validation, conditional formatting, auto filter | ✅ | |
| Data | pivot tables | ⚠️ | Programmatic creation exists; editing existing pivots is not modeled, but round-trip preservation is supported. |
| Strings | shared strings, rich text runs | ✅ | `XSSFRichTextString` and `<rPr>` support are present. |
| Other | workbook/sheet protection, xlsm macro preservation | ✅ | VBA bytes are preserved in macro-enabled round-trips. |
| Other | sparklines | ❌ | |
| Other | external data connections | 🔵 | `xl/connections.xml` and `xl/externalLinks/*` round-trip via `_preservedEntries`. |

#### docx / XWPF (~65%)

| Category | Feature | Status | Notes |
|---|---|---|---|
| Paragraphs/Runs | text read/write | ✅ | |
| 〃 | font (bold/italic/underline/strikeout/name/size/color) | ✅ | Round-trip covered. |
| 〃 | alignment (left/center/right/both) | ✅ | |
| 〃 | indentation (left/right/firstLine/hanging) | ✅ | |
| 〃 | spacing (before/after/line) | ✅ | |
| 〃 | bullet/numbered lists | ✅ | OOXML numbering is implemented. |
| Tables | create/read tables, rows, cells | ✅ | Round-trip covered. |
| 〃 | cell merge and table borders | 🔵 | Existing merge/borders preserved via raw XML; API-level creation not modeled. |
| Sections | page setup (size/margins/orientation) | ✅ | |
| 〃 | headers/footers | ✅ | Rich content (images, formatting) in headers/footers preserved via `_preservedEntries` when not modified via API. |
| 〃 | columns | ✅ | `setColumns()`/`getColumnCount()`/`getColumnSpacing()` API, round-trip verified |
| Links | external hyperlinks | ✅ | |
| Images | inline images and rotation | ✅ | |
| 〃 | floating (anchored) images | 🔵 | `<wp:anchor>` elements preserved via raw XML capture/re-emission. |
| 〃 | text boxes (`w:txbxContent`) | ❌ | Text inside Word text boxes is not read. |
| Review | comments | 🔵 | Existing comments round-trip via `_preservedEntries`; API creation/editing not modeled. |
| 〃 | footnotes/endnotes | 🔵 | Existing parts round-trip via `_preservedEntries`; API creation/editing not modeled. |
| Fields | TOC, page numbers, mail merge-style fields | ✅ | Write/read/round-trip covered. |
| Content Controls | SDT (structured document tags) | 🔵 | Block-level and inline SDT preserved via raw XML capture/re-emission. |
| Styles | paragraph, character, and table styles | ❌ | Direct formatting is supported; style model parity is not. |
| Track Changes | insertions/deletions/moves | ❌ | |
| Other | docm macro preservation | ✅ | VBA byte preservation. |
| 〃 | unknown part preservation | ✅ | `_preservedEntries` mechanism implemented. |
| 〃 | OLE embeddings | 🔵 | `word/embeddings/*` round-trip via `_preservedEntries`. |

#### pptx / XSLF (~40%)

Simple presentation creation and editing is usable: create/read slides, text boxes, formatted runs, pictures, rotation, tables, and slide size are covered, with Java POI interop tests for basic generated presentations. More advanced PowerPoint features such as charts, SmartArt, notes, media, layouts, masters, themes, animations, and grouped shapes are mostly preserved during round-trip rather than exposed as editable object models.

| Category | Feature | Status | Notes |
|---|---|---|---|
| Slides | create/read slides, slide size | ✅ | |
| 〃 | notes slides | 🔵 | Existing notes slide parts round-trip via `_preservedEntries`; API-level creation/editing is not modeled. |
| Text | text boxes, multiple paragraphs, run formatting | ✅ | Bold, italic, underline, strikeout, size, font, and color are covered. |
| Shapes | pictures, anchors, size, rotation | ✅ | Round-trip covered. |
| 〃 | tables | ✅ | `p:graphicFrame` / `a:tbl` write/read is implemented. |
| 〃 | group shapes, connectors | 🔵 | Unknown `p:spTree` children preserved verbatim via raw XML capture/re-emission. |
| 〃 | SmartArt, charts | 🔵 | Existing parts are preserved, but not modeled. |
| Media | video/audio embedding | 🔵 | Non-image `ppt/media/*` parts round-trip via `_preservedEntries`; API-level embedding is not modeled. |
| Animation | animations and transitions | 🔵 | Preserved as unknown parts where present. |
| Theme | layouts, masters, themes | 🔵 | Preserved, not editable. |
| Other | pptm macro preservation, unknown part preservation | ✅ | |

#### xls / HSSF (~10%)

| Category | Feature | Status | Notes |
|---|---|---|---|
| Basic workbook | simple read/write of cell values | ⚠️ in progress | Current Phase 4 work. |
| BIFF/OLE2 | broader records, styles, formulas, images, charts, filters, pivots | ❌ | Legacy format work remains intentionally narrow for now. |

#### Legacy binary formats

| Format | Status | Notes |
|---|---|---|
| doc / HWPF | ~5% | Read-only MVP/stub level. |
| ppt / HSLF | ~5% | Read-only MVP/stub level. |

### Practical Gaps

Highest priority gaps:

| # | Gap | Formats | Why it matters |
|---|---|---|---|
| 1 | Formula evaluation | xlsx | Template fill → save → open in Excel works, but programmatic access to newly calculated results needs a real evaluator. |
| 2 | Chart creation | xlsx, pptx | Existing charts can be preserved, but report/presentation generation often needs to create charts from data. |
| 3 | Comment API model | xlsx, docx | Existing comments survive round-trip through `_preservedEntries`, but API-level read/create/edit support is not implemented. |
| 4 | HSSF `.xls` depth | xls | Basic values exist, but full BIFF record, style, formula, and drawing support is still in progress. |
| 5 | docx styles and text boxes | docx | Word documents commonly rely on styles and text boxes for layout and semantics. SDT is now preserved. |

Lower priority gaps include SmartArt, animations, transitions, `xls`/`doc`/`ppt` legacy depth, tracked changes, and sparklines.

### Test Coverage Snapshot

Tracked in [NOW.md](./NOW.md):

| Project | Tests | Notes |
|---|---:|---|
| Core.Tests | 244 | +3 POI test-file preservation verification tests, +1 xlsx auto-shape, +1 docx header/footer preservation, +2 docx SDT preservation, +1 docx floating anchor images preservation. |
| Formula.Tests | 10 | Minimal formula package coverage. |
| Interop.Tests (C#) | 55 | Bidirectional Java/.NET fixtures and preservation tests. |
| **Total (C#)** | **309** | |
| Java POI side (Maven) | 44 | Java fixture generation/readback tests. |

---

## Quick Start

Published to **[NuGet.org](https://www.nuget.org/packages/DotnetPoi.Core)** — `dotnet add package DotnetPoi.Core`.

```bash
git clone --recurse-submodules https://github.com/modeverv/dotnet-poi
cd dotnet-poi
dotnet build
dotnet test
```

### Usage

```csharp
using DotnetPoi.XSSF.UserModel;

var workbook = new XSSFWorkbook();
var sheet = workbook.createSheet("Sheet1");
var row = sheet.createRow(0);
row.createCell(0).setCellValue("Hello");
row.createCell(1).setCellValue(42);

using var fs = new FileStream("output.xlsx", FileMode.Create);
workbook.write(fs);
```

Runnable examples:

```bash
dotnet run --project examples/Phase0WriteExample/Phase0WriteExample.csproj
dotnet run --project examples/Phase1InteropExample/Phase1InteropExample.csproj
dotnet run --project examples/Phase25ImagesExample/Phase25ImagesExample.csproj
dotnet run --project examples/Phase32DocxExample/Phase32DocxExample.csproj
dotnet run --project examples/Phase33PptxExample/Phase33PptxExample.csproj
dotnet run --project examples/Phase34AgileEncryptionExample/Phase34AgileEncryptionExample.csproj
dotnet run --project examples/Phase4HssfXlsExample/Phase4HssfXlsExample.csproj
dotnet run --project examples/Phase5FormulaEvaluatorExample/Phase5FormulaEvaluatorExample.csproj
dotnet run --project examples/Phase7CellTypesExample/Phase7CellTypesExample.csproj
dotnet run --project examples/Phase8CoreOnlyExample/Phase8CoreOnlyExample.csproj
dotnet run --project examples/UsageSamples/UsageSamples.csproj
```

pptx example:

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

docx example:

```csharp
using DotnetPoi.XWPF.UserModel;

using var doc = new XWPFDocument();

var para = doc.createParagraph();
var run = para.createRun();
run.setText("Hello from dotnet-poi");
run.setBold(true);

using var fs = new FileStream("output.docx", FileMode.Create);
doc.write(fs);
```

---

## Why This Project

The .NET Excel library landscape has structural problems:

- **NPOI**: Supports both xls and xlsx, but v2.8.0+ requires a commercial maintenance fee
- **ClosedXML / EPPlus**: xlsx only — cannot handle xls (BIFF format)

dotnet-poi aims to solve both problems by porting Apache POI — a battle-tested implementation — transparently and faithfully, with **no licensing strings attached, ever**.

---

## Porting Approach

Apache POI source is kept as a git submodule under `poi/`, so the original Java is always at hand. LLMs handle the mechanical Java → C# conversion; humans handle architecture decisions and quality verification.

This project is also an experiment: **can LLMs carry a large-scale, long-running intellectual porting effort?**

See [agents.md](./agents.md) for detailed porting rules.

---

## Repository Structure

```
dotnet-poi/
├── .github/
│   ├── workflows/              # CI, examples, XML parity fixture workflows
│   └── java-upgrade/           # Java upgrade helper hooks/scripts
├── poi/                        # Apache POI submodule (read-only reference)
├── src/
│   ├── DotnetPoi.Core/         # NuGet: DotnetPoi.Core
│   │   ├── SS/                 # Common interfaces, usermodel types, XML writer
│   │   ├── POIFS/              # OLE2 / CFB container and Agile encryption support
│   │   ├── XSSF/               # xlsx / xlsm
│   │   ├── HSSF/               # xls / BIFF bootstrap
│   │   ├── XWPF/               # docx / docm
│   │   ├── HWPF/               # doc read-only MVP
│   │   ├── XSLF/               # pptx / pptm
│   │   └── HSLF/               # ppt read-only MVP
│   └── DotnetPoi.Formula/      # NuGet: DotnetPoi.Formula
│       └── UserModel/          # FormulaEvaluator implementation
├── tests/
│   ├── DotnetPoi.Core.Tests/       # Core tests grouped by module
│   ├── DotnetPoi.Formula.Tests/    # Formula package tests
│   ├── DotnetPoi.Interop.Tests/    # Java/.NET compatibility tests
│   │   ├── java/                   # Maven project using Apache POI
│   │   └── fixtures/               # from-poi, from-dotnet-poi, XML parity, preservation fixtures
│   └── test-files/                 # Shared binary fixtures (xlsm, docm, pptm, images)
├── examples/
│   ├── UsageSamples/               # Current user-facing sample set
│   ├── Phase*Example/              # Historical phase/progress examples
│   ├── EdgeCaseProbeExample/       # Ad hoc edge-case probe sample
│   ├── README.md
│   └── output/                     # Generated example outputs
├── docs_src/
│   ├── site.json                   # Docs nav and site metadata
│   ├── content/                    # Source Markdown
│   ├── assets/                     # Source docs assets
│   └── templates/                  # Docs generator templates
├── docs/                           # Generated static documentation site
├── tools/
│   ├── DotnetPoi.DocsGenerator/    # docs_src -> docs generator
│   ├── XmlCheck/                   # XML inspection/check helper
│   ├── dev/                        # Docker devbox compose/env/Dockerfile
│   ├── porter/                     # Porting progress tracker
│   └── test.sh                     # Local interop test runner
├── debugtest/                      # Local scratch/debug project
├── DotnetPOI.sln                   # Main solution
├── global.json                     # .NET SDK pin
├── NOW.md                          # Current coverage snapshot
├── CHECKPOINT.md                   # Working notes / handoff log
├── agents.md                       # LLM agent instructions
├── README.jp.md                    # Japanese README
├── README.save.md                  # Saved README copy
├── POI_INTEGRATION_FIXTURE_TODO.md
├── XMLBEANS_XML_OUTPUT_TODO.md
└── README.md
```

> **Architecture note:** All format implementations live under `DotnetPoi.Core/` as a single assembly.
> The old per-format project directories (`DotnetPoi.XSSF/`, `DotnetPoi.SS/`, etc.) no longer
> exist as separate `.csproj` files — their source files have been consolidated into Core.
> Namespaces are unchanged (e.g. `DotnetPoi.XSSF.UserModel`).

---

## Contributing

This is a personal long-term project, but PRs and Issues are welcome. Please read [agents.md](./agents.md) before contributing.

---

## License

[Apache License 2.0](./LICENSE) — same as upstream Apache POI.

---

## Disclaimer

This project is not affiliated with the Apache Software Foundation or the Apache POI project. Apache POI is a registered trademark of the Apache Software Foundation.
