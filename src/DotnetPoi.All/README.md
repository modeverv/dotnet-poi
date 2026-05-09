# DotnetPoi.All

An **unofficial**, faithful port of [Apache POI](https://poi.apache.org/) for .NET.

This is the meta-package that includes all `dotnet-poi` functionality in a single dependency. It references:
- `DotnetPoi.Ooxml` (stable)
- `DotnetPoi.Legacy` (in-development)
- `DotnetPoi.Formula` (narrow subset)
- `DotnetPoi.Common` (transitive)
- `DotnetPoi.POIFS` (transitive)

For most users, this is the recommended package.

## Install

Simple presentation creation and editing is usable: create/read slides, text boxes, formatted runs, pictures, rotation, tables, and slide size are covered, with Java POI interop tests for basic generated presentations. More advanced PowerPoint features such as charts, SmartArt, notes, media, layouts, masters, themes, animations, and grouped shapes are mostly preserved during round-trip rather than exposed as editable object models.

```shell
dotnet add package DotnetPoi.All
```

```xml
<PackageReference Include="DotnetPoi.All" Version="..." />
```

## Document

Document is [dotnet-poi Documentation](https://modeverv.github.io/dotnet-poi/) 

## Current scope

- OOXML (`xlsx`, `docx`, `pptx`) is the most practical surface. `docx` tracked-change XML is preserved during round-trip, but revision accept/reject/create/edit APIs are not modeled.
- Legacy binary formats are included for convenience, but remain partial: `xls` and `doc` cover practical read/write/light-edit slices, while `ppt` is preservation/text-extraction oriented.
- Formula evaluation is intentionally limited. Formula text and cached-value preservation live in the format packages; full Excel-compatible calculation is not a project goal.

---

## Philosophy

- 🔱 Maximum fidelity to upstream Apache POI — we follow, not reinvent
- 🤖 Ported class-by-class with LLM assistance, with tests written alongside
- 💸 Free forever. No EULA. No maintenance fee. No exceptions.
- 📖 Apache POI is the source of truth — included as a git submodule
- ⚠️ Not affiliated with the Apache Software Foundation

---

## Why This Project

The .NET Excel library landscape has structural problems:

- **NPOI**: Supports both xls and xlsx, but v2.8.0+ requires a commercial maintenance fee
- **ClosedXML / EPPlus**: xlsx only — cannot handle xls (BIFF format)

dotnet-poi aims to solve both problems by porting Apache POI — a battle-tested implementation — transparently and faithfully, with **no licensing strings attached, ever**.

---

## Format Coverage

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
| Use this if you… | Otherwise consider… |
|---|---|
| Want everything with one dependency | `DotnetPoi.Ooxml` for OOXML-only projects |
| Don't want to think about granular package selection | `DotnetPoi.Legacy` for legacy-only projects |
| Need all formats + formula evaluation | `DotnetPoi.Ooxml` + `DotnetPoi.Formula` for minimal footprint |

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

---

### Test Coverage Snapshot

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

## License

[Apache License 2.0](../../LICENSE) — same as upstream Apache POI.

## Disclaimer

This project is not affiliated with the Apache Software Foundation or the Apache POI project.
Apache POI is a registered trademark of the Apache Software Foundation.

