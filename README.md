# dotnet-poi

An **unofficial**, faithful port of [Apache POI](https://poi.apache.org/) for .NET.

[![CI](https://github.com/modeverv/dotnet-poi/actions/workflows/ci.yml/badge.svg)](https://github.com/modeverv/dotnet-poi/actions/workflows/ci.yml)
[![Examples](https://github.com/modeverv/dotnet-poi/actions/workflows/examples.yml/badge.svg)](https://github.com/modeverv/dotnet-poi/actions/workflows/examples.yml)
[![XML Parity](https://github.com/modeverv/dotnet-poi/actions/workflows/xml-parity-fixtures.yml/badge.svg)](https://github.com/modeverv/dotnet-poi/actions/workflows/xml-parity-fixtures.yml)
[![NuGet All](https://img.shields.io/nuget/v/DotnetPoi.All)](https://www.nuget.org/packages/DotnetPoi.All)
[![NuGet Ooxml](https://img.shields.io/nuget/v/DotnetPoi.Ooxml)](https://www.nuget.org/packages/DotnetPoi.Ooxml)
[![NuGet Formula](https://img.shields.io/nuget/v/DotnetPoi.Formula)](https://www.nuget.org/packages/DotnetPoi.Formula)
[![NuGet Legacy](https://img.shields.io/nuget/v/DotnetPoi.Legacy)](https://www.nuget.org/packages/DotnetPoi.Legacy)
![License](https://img.shields.io/badge/license-Apache%202.0-blue)
![Status](https://img.shields.io/badge/status-practical%20OOXML%20workflows-brightgreen)
![Tests](https://img.shields.io/badge/tests-550%20passing-brightgreen)

## NuGet Package Strategy

dotnet-poi uses a **multi-package architecture** with clear separation of concerns. This enables OOXML stability to advance independently of legacy binary format development.

### Recommended Package

For most users, **`DotnetPoi.All`** is the simplest choice — it includes the stable OOXML package plus the current Legacy and Formula packages with a single dependency:

```xml
<PackageReference Include="DotnetPoi.All" Version="1.0.1" />
```

### Dependency Picker

Decide what you need, then install the matching packages:

| Scenario | Packages to install | Notes |
|---|---|---|
| **OOXML** (xlsx / docx / pptx) <br/>read/write | `DotnetPoi.Ooxml` | Modern Office 2007+ formats only — minimal footprint |
| **OOXML** + formula evaluator | `DotnetPoi.Ooxml`<br/>`DotnetPoi.Formula` | Add `createFormulaEvaluator()` support when needed |
| **Legacy binary** (xls / doc / ppt) <br/>read/write | `DotnetPoi.Legacy` | BIFF-based Office 97-2003 formats |
| **Legacy binary** + formula evaluator | `DotnetPoi.Legacy`<br/>`DotnetPoi.Formula` | |
| **All formats**, no formula evaluator | `DotnetPoi.Ooxml`<br/>`DotnetPoi.Legacy` | Two packages, excludes formula engine (smaller attack surface) |
| **Everything** (all formats + formula) | `DotnetPoi.All` | One dependency. OOXML is the stable 1.0 surface; Legacy and Formula remain partial. |

Transitive dependencies (`DotnetPoi.Common`, `DotnetPoi.POIFS`) are resolved automatically by NuGet.

### Granular Packages

| Package | Contents | Best for |
|---|---|---|
| **DotnetPoi.Ooxml** | XSSF (xlsx/xlsm), XWPF (docx/docm), XSLF (pptx/pptm) + OPC/openxml package + shared POIFS foundation | Users who work with modern Office 2007+ formats only |
| **DotnetPoi.Legacy** | HSSF (xls), HWPF (doc), HSLF (ppt) | Users who need legacy binary format support |
| **DotnetPoi.Formula** | Formula evaluator (`IFormulaEvaluator`, `FormulaEvaluator`, `CellValue`) | Only when you need the supported formula evaluator subset |
| **DotnetPoi.Common** | SS interfaces, shared enums, common exceptions, XML writer foundation | Base dependency (included transitively by all packages above) |
| **DotnetPoi.POIFS** | OLE2/CFB compound file container, encryption helpers | OLE2 container foundation (included transitively by Ooxml and Legacy) |
| **DotnetPoi.All** | Stable OOXML 1.0 plus the current Legacy, Formula, Common, and POIFS packages | Users who want everything with one dependency |

### Migration from `DotnetPoi.Core`

The legacy `DotnetPoi.Core` facade package has been **removed**. Replace any existing `DotnetPoi.Core` reference with `DotnetPoi.All` — namespaces and public API surface are unchanged:

```xml
<!-- Before -->
<PackageReference Include="DotnetPoi.Core" Version="0.5.0" />

<!-- After -->
<PackageReference Include="DotnetPoi.All" Version="1.0.1" />
```

### Design Principle

All format packages have **zero knowledge of `Formula`**. Adding `DotnetPoi.Formula` to your project automatically enables `createFormulaEvaluator()` via lazy assembly discovery at runtime. Without it, the call throws a clear `NotSupportedException`.

```xml
<!-- OOXML-only projects: no formula engine pulled in -->
<PackageReference Include="DotnetPoi.Ooxml" Version="..." />

<!-- Add formula evaluation when needed -->
<PackageReference Include="DotnetPoi.Ooxml" Version="..." />
<PackageReference Include="DotnetPoi.Formula" Version="..." />
```

**Why this split:**
- **Ooxml can be stable early** — all OOXML read/write/format logic is self-contained. No dependency on legacy binary development.
- **Legacy can evolve safely** — HSSF/HWPF/HSLF development can proceed without destabilizing OOXML users.
- **Formula stays narrow** — full Excel-compatible evaluation is not a current project goal and can remain separate.
- **Smaller dependency for simple use cases** — users who only need xlsx don't pull in legacy format code.
- **Security** — applications handling untrusted documents can omit the formula evaluator entirely, reducing the attack surface.

## Philosophy

- 🔱 Maximum fidelity to upstream Apache POI — we follow, not reinvent
- 🤖 Ported class-by-class with LLM assistance, with tests written alongside
- 💸 Free forever. No EULA. No maintenance fee. No exceptions.
- 📖 Apache POI is the source of truth — included as a git submodule
- ⚠️ Not affiliated with the Apache Software Foundation

---

## Status

Current status: **1.0.x for covered OOXML workflows** — packages are available on [NuGet.org](https://www.nuget.org/).

Version 1.0 means the documented OOXML workflows are treated as stable. It does **not** mean full Apache POI parity or complete Office feature coverage.

| Package | NuGet ID | Version | Status |
|---|---|---------|---|
| **All** | `DotnetPoi.All` | 1.0.x   | Meta-package: OOXML 1.0 plus partial Legacy and Formula packages |
| **OOXML** | `DotnetPoi.Ooxml` | 1.0.x   | Stable for common xlsx/docx/pptx workflows |
| **Common** | `DotnetPoi.Common` | 1.0.x   | Shared API/support package, pulled transitively |
| **POIFS** | `DotnetPoi.POIFS` | 1.0.x   | OLE2/CFB support package, pulled transitively |
| **Legacy** | `DotnetPoi.Legacy` | 0.5.x   | In-development (HSSF/HWPF/HSLF) |
| **Formula** | `DotnetPoi.Formula` | 0.1.x   | Narrow evaluator subset |

The strongest format today is **xlsx / XSSF**, with broad support for workbook creation, reading, editing, styling, layout, images, formulas-as-text, macro preservation, and Java POI interop. **docx / XWPF** and **pptx / XSLF** are also useful for practical generation, light editing, and loss-resistant round-trips of many real files.

This does **not** mean the whole Apache POI surface is complete. Advanced OOXML features such as chart creation and docx comment editing are still limited, some features are preservation-only rather than modeled APIs, and formula evaluation remains intentionally narrow. Legacy binary formats have improved: `.xls` now has practical basic workbook read/write, styling/layout slices, preservation, and Java POI interop coverage; `.doc` can extract body text and perform limited body edits with preservation. `.ppt` is still early. In short: **use it today for the supported workflows shown below; check the matrix before relying on an advanced or legacy feature.**

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
| Drawings | charts | 🔵 | Existing chart parts are preserved during round-trip; new creation/editing is not modeled. |
| Review | comments | ✅ | Cell comment read/create/edit/remove is modeled via `XSSFComment`, cell/sheet lookup, and VML/comment part write/read. Rich formatting and VML shape styling are still minimal. |
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
| 〃 | text boxes (`w:txbxContent`) | ✅ | Text extraction from inline and anchored drawing textboxes is supported. |
| Review | comments | 🔵 | Existing comments round-trip via `_preservedEntries`; API creation/editing not modeled. |
| 〃 | footnotes/endnotes | 🔵 | Existing parts round-trip via `_preservedEntries`; API creation/editing not modeled. |
| Fields | TOC, page numbers, mail merge-style fields | ✅ | Write/read/round-trip covered. |
| Content Controls | SDT (structured document tags) | 🔵 | Block-level and inline SDT preserved via raw XML capture/re-emission. |
| Styles | paragraph style reference (pStyle) | ✅ | `setStyle()`/`getStyleID()` API, round-trip verified. Character/table styles ❌. `word/styles.xml` 🔵 preserved + default styles auto-generated for new docs. |
| Track Changes | insertions/deletions/moves | 🔵 | Tracked-change XML is preserved in body/paragraph child order during round-trip; API-level accept/reject/create/edit is not modeled. |
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

#### xls / HSSF (~35%)

| Category | Feature | Status | Notes |
|---|---|---|---|
| Cell values | string, numeric, boolean, blank, error | ✅ | BIFF8 LabelSST/Number/BoolErr/Blank round-trip covered. |
| Sheets | multiple sheets, sparse rows/cells, high column indexes | ✅ | |
| Styles | fonts, data formats, alignment, wrap, borders, fills | ⚠️ | Core HSSFFont/HSSFCellStyle round-trip works for common cases; not full BIFF style parity. |
| Layout | column width, row height, hidden rows/columns, merged regions, freeze panes | ✅ | |
| Formulas | formula text + cached value read | ⚠️ | Existing POI formula fixtures can be read; new BIFF formula token writing and evaluation are not implemented. |
| Compatibility | representative POI `.xls` fixture loading | ✅ | Includes basic, styles, formulas, hyperlinks, comments, drawings, images, and macro fixtures as load/preservation cases. |
| Interop | Java POI bidirectional fixtures | ⚠️ | basic/styles/layout/unicode/comprehensive fixture coverage. |
| Preservation | non-Workbook OLE streams, VBA streams, unknown BIFF records | ✅ | Light edits preserve unmodeled streams/records where possible. |
| Not modeled | images/shapes/charts/comments/hyperlink editing/filters/pivots | ❌ | Some are load/preservation fixtures, but not public usermodel creation/edit APIs. |

#### doc / HWPF (~25%)

| Category | Feature | Status | Notes |
|---|---|---|---|
| Reading | OLE2 `.doc` open, FIB/table stream parsing | ✅ | `WordDocument` + `0Table`/`1Table` selection and fallback covered. |
| Text | main body text extraction | ✅ | CLX/piece table based extraction with compressed and Unicode text pieces. |
| UserModel | Range, Paragraph, CharacterRun | ⚠️ | Paragraph/run splitting and some offsets/composition covered. |
| Formatting | character and paragraph properties | ⚠️ | CHPX-derived font name/size/bold/italic/underline/strike and minimal PAPX fields. |
| Extraction | **header/footer and table structures** | ✅ | `getHeaderStoryRange()`, table row/cell iteration implemented. |
| Editing | no-op write, append paragraph, simple text replacement | ⚠️ | Limited main-body edit path; not a full Word binary editing engine. |
| Preservation | OLE streams/storages, embedded OLE | ✅ | Unedited stream/storage content is preserved in representative fixtures. |
| Interop | Java POI bidirectional testing | ⚠️ | Java POI correctly extracts tables and header/footer text from dotnet-poi saved files. |
| Not modeled | images/footnotes/comments/fields API | ❌ | Streams may be preserved, but these are not usermodel creation/edit features. |

#### ppt / HSLF (~5%)

| Category | Feature | Status | Notes |
|---|---|---|---|
| Reading | open OLE2 `.ppt` and scan `PowerPoint Document` records | ⚠️ | Minimal reader exists. |
| Text | slide text extraction from TextChars/TextBytes atoms | ⚠️ | Early recursive scan; slide order/persist-pointer fidelity is still planned. |
| Writing | no-op preservation / editing | ❌ | Planned next: HSLF stream inventory, no-op write, Java POI interop. |

### Practical Gaps

Highest priority gaps:

| # | Gap | Formats | Why it matters |
|---|---|---|---|
| 1 | Full formula evaluation | xlsx | Template fill → save → open in Excel works; programmatic access to newly calculated results beyond the small DotnetPoi.Formula subset needs a real calculation engine. |
| 2 | Chart creation | xlsx, pptx | Existing charts can be preserved, but report/presentation generation often needs to create charts from data. |
| 3 | Comment API depth | docx | Existing docx comments survive round-trip and can be read through minimal `XWPFComment` APIs. Minimal create/edit is available through `createComment(...)`, mutable comment metadata/text, and paragraph range marker insertion; richer comment content and cleanup-heavy editing remain limited. xlsx cell comments are already modeled for common read/create/edit/remove workflows. |
| 4 | HSSF/HWPF depth | xls, doc | Basic legacy read/write and preservation exist, but images, shapes, advanced formatting, and complete editing are still limited. |
| 5 | docx style depth and revision APIs | docx | Paragraph style references and tracked-change preservation are supported, but full character/table style editing and accept/reject/create/edit APIs for revisions remain limited. |

Lower priority gaps include SmartArt, animations, transitions, `ppt` legacy depth, tracked-change editing APIs, and sparklines.

### Test Coverage Snapshot

Tracked in [NOW.md](./NOW.md):

| Package | Test Project | Tests | Notes |
|---|---|---|---:|
| **OOXML** | Ooxml.Tests | 169 | OOXML-specific split tests |
| **Legacy** | Legacy.Tests | 224 | Legacy-specific split tests |
| **Formula** | Formula.Tests | 11 | Minimal formula package coverage |
| **Common** | Common.Tests | 79 | Shared SS/utility tests |
| **POIFS** | POIFS.Tests | 11 | OLE2 container tests |
| **All** | All.Tests | 7 | Meta-package smoke tests |
| **Interop** | Interop.Tests (C#) | 71 passed / 2 skipped | Bidirectional Java/.NET fixtures + preservation |
| **Total (C#)** | | **572 passed / 2 skipped** | |
| Java POI side (Maven) | | 45+ | Java fixture generation/readback tests |

---

## Testing Strategy

This project employs a multi-layered testing strategy to ensure maximum fidelity to Apache POI and seamless interoperability with Microsoft Office.

- **Unit Tests (xUnit):** Ported alongside each class from the original Apache POI test suite. Ensures internal logic and edge-case handling are consistent with Java.
- **XML Parity Tests:** We verify that our `PoiXmlWriter` produces byte-equivalent XML output to Apache POI (XMLBeans). This ensures that subtle formatting differences don't break digital signatures or strict OOXML parsers.
- **Bidirectional Interop Tests:** Every supported format is tested in both directions:
    - **Direction A:** Java POI writes → dotnet-poi reads.
    - **Direction B:** dotnet-poi writes → Java POI reads.
- **Preservation Tests:** We verify that unmodeled features (macros, charts, comments, pivot tables) survive a read-modify-write cycle (round-trip) without data loss or corruption.
- **Release Hygiene:** CI packs `DotnetPoi.Common` → `DotnetPoi.POIFS` → `DotnetPoi.Legacy` → `DotnetPoi.Formula` → `DotnetPoi.Ooxml` → `DotnetPoi.All`, validates tag/package metadata, checks package READMEs, and installs every local nupkg from a temporary NuGet source before publish.
- **Manual Verification:** Before releases, we perform manual checks using real **Microsoft Office (Excel/Word/PowerPoint)** and **LibreOffice** on macOS, Windows, and Linux to ensure no "repair" dialogs or visual regressions occur.

---

## Quick Start

Published to **[NuGet.org](https://www.nuget.org/packages/DotnetPoi.All)** — `dotnet add package DotnetPoi.All`.

or

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
│   ├── DotnetPoi.Common/       # Shared interfaces, enums, utilities, XML writer
│   ├── DotnetPoi.POIFS/        # OLE2 / CFB container and Agile encryption support
│   ├── DotnetPoi.Ooxml/        # XSSF (xlsx/xlsm), XWPF (docx/docm), XSLF (pptx/pptm)
│   │   ├── XSSF/               # xlsx / xlsm
│   │   ├── XWPF/               # docx / docm
│   │   └── XSLF/               # pptx / pptm
│   ├── DotnetPoi.Legacy/       # HSSF (xls), HWPF (doc), HSLF (ppt)
│   │   ├── HSSF/               # xls / BIFF basic workbook + preservation
│   │   ├── HWPF/               # doc text extraction + limited body editing
│   │   └── HSLF/               # ppt minimal reader
│   ├── DotnetPoi.Formula/      # NuGet: DotnetPoi.Formula
│   │   └── UserModel/          # FormulaEvaluator implementation
│   └── DotnetPoi.All/          # Meta-package referencing everything
├── tests/
│   ├── DotnetPoi.Common.Tests/     # Common package tests
│   ├── DotnetPoi.POIFS.Tests/      # POIFS container tests
│   ├── DotnetPoi.Ooxml.Tests/      # OOXML format tests
│   ├── DotnetPoi.Legacy.Tests/     # Legacy format tests
│   ├── DotnetPoi.All.Tests/        # All-package smoke tests
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
│   ├── release/                    # package hygiene and NuGet install smoke scripts
│   ├── porter/                     # Porting progress tracker
│   └── test.sh                     # Local interop test runner
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

> **Architecture note:** Format implementations are split across `DotnetPoi.Ooxml` (OOXML formats) and `DotnetPoi.Legacy` (legacy binary formats), with `DotnetPoi.Common` and `DotnetPoi.POIFS` as shared foundations. `DotnetPoi.All` is the meta-package that bundles everything under a single dependency.

---

## Contributing

This is a personal long-term project, but PRs and Issues are welcome. Please read [agents.md](./agents.md) before contributing.

---

## License

[Apache License 2.0](./LICENSE) — same as upstream Apache POI.

---

## Disclaimer

This project is not affiliated with the Apache Software Foundation or the Apache POI project. Apache POI is a registered trademark of the Apache Software Foundation.
