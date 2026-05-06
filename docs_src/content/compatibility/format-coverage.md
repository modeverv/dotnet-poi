# Format Coverage

Legend: **✅** complete / **⚠️** partial (write-only, etc.) / **🔵** preserved as unknown parts, not editable / **❌** not implemented / **—** not applicable

## xlsx / XSSF (~78%)

| Category | Feature | Status | Notes |
|---|---|---|---|
| Cell values | string, numeric, date, boolean, error | ✅ | |
| Formulas | formula text write/read + cached value | ✅ | Excel recalculation-on-open works |
| Formulas | full formula evaluation | ❌ deferred | See package-split |
| Styles | fonts, fills, borders, number formats, alignment | ✅ | Round-trip verified |
| Layout | merged cells, column width, row height | ✅ | |
| Layout | hidden rows/columns, freeze panes | ✅ | |
| Layout | active cell, sheet selection, active sheet | ✅ | Active cell/selected in-memory; active sheet round-trips |
| Layout | print settings (margins, paper size, orientation, headers) | ✅ | |
| Drawings | images, anchors, rotation, hyperlinks | ✅ | |
| Drawings | charts, comments | 🔵 | Preserved on round-trip, new creation not modeled |
| Drawings | auto-shapes | ❌ | |
| Data | data validation, conditional formatting, auto filter | ✅ | |
| Data | pivot tables | ⚠️ | Programmatic creation works; editing existing not modeled |
| Strings | shared strings, rich text runs | ✅ | Per-character formatting via XSSFRichTextString |
| Protection | workbook/sheet protection | ✅ | |
| Macros | xlsm preservation | ✅ | VBA bytes preserved on round-trip |
| Other | sparklines | ❌ | |
| Other | external data connections | 🔵 | `xl/connections.xml` / `xl/externalLinks/*` round-trip via `_preservedEntries` |

## docx / XWPF (~65%)

| Category | Feature | Status | Notes |
|---|---|---|---|
| Paragraphs/runs | text, font name/size/color, bold, italic, underline, strikeout | ✅ | Round-trip verified |
| Paragraphs | alignment, indents, spacing, bullet/numbered lists | ✅ | |
| Tables | create/read tables, rows, cells | ✅ | Round-trip verified |
| Tables | cell merging, borders | ❌ | |
| Sections | page size, margins, orientation | ✅ | |
| Sections | headers and footers | ✅ | Round-trip verified |
| Sections | columns | ❌ | |
| Links | hyperlinks (external URLs) | ✅ | |
| Images | inline images with rotation | ✅ | |
| Images | text boxes (w:txbxContent) | ❌ | |
| Annotations | comments | 🔵 | Existing parts round-trip via `_preservedEntries` |
| Annotations | footnotes, endnotes | 🔵 | `word/footnotes.xml` / `word/endnotes.xml` round-trip preserved |
| Fields | TOC, page numbers, mail merge | ✅ | Write/read/round-trip |
| SDT | content controls | ❌ | |
| Styles | paragraph/character/table styles | ❌ | Direct formatting only; `word/styles.xml` 🔵 preserved but style refs in document.xml lost on model rewrite |
| Track Changes | revision marks | ❌ | |
| Other | OLE embeddings | 🔵 | `word/embeddings/*` round-trip via `_preservedEntries` |
| Other | docm macro preservation | ✅ | VBA bytes preserved |
| Other | unknown part preservation | ✅ | `_preservedEntries` mechanism preserves non-model ZIP entries |

## pptx / XSLF (~40%)

| Category | Feature | Status | Notes |
|---|---|---|---|
| Slides | create/read slides | ✅ | |
| Slides | slide size | ✅ | |
| Slides | notes slides | 🔵 | `ppt/notesSlides/notesSlide*.xml` round-trip via `_preservedEntries` |
| Text | text boxes (p:sp) create/read | ✅ | Round-trip verified |
| Text | multiple paragraphs, run formatting | ✅ | bold, italic, underline, size, font, color |
| Shapes | images with position, size, rotation | ✅ | Round-trip verified |
| Shapes | tables (p:graphicFrame/a:tbl) | ✅ | Round-trip verified |
| Shapes | grouping, connectors, lines | ❌ | |
| Shapes | SmartArt, charts | 🔵 | Preserved as unknown parts |
| Media | video/audio embedding | 🔵 | Non-image `ppt/media/*` round-trip via `_preservedEntries` |
| Animation | animations, transitions | 🔵 | Preserved as unknown parts |
| Theme | layout, master, theme | 🔵 | Not editable, preserved on round-trip |
| Other | pptm macro preservation | ✅ | VBA bytes preserved |

## xls / HSSF (~10%)

| Category | Feature | Status | Notes |
|---|---|---|---|
| General | basic value read/write | ⚠️ | 2 tests only |
| General | everything else (BIFF records, styles, images, charts, formulas, filters, pivots) | ❌ | Legacy format, low priority |

## Legacy Formats

| Format | Status | Notes |
|---|---|---|
| doc (HWPF) | ~5% | Read stub only |
| ppt (HSLF) | ~5% | Read stub only |
