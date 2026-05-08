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
| Drawings | charts | 🔵 | Existing chart parts are preserved on round-trip; new creation/editing is not modeled |
| Review | comments | ✅ | Cell comment read/create/edit/remove is modeled via `XSSFComment`, cell/sheet lookup, and VML/comment part write/read. Rich formatting and VML shape styling are still minimal |
| Drawings | auto-shapes, group shapes, connectors | 🔵 | Unknown `xdr:twoCellAnchor` children in drawing.xml preserved verbatim via raw XML capture/re-emission. Currently only `xdr:pic` is modeled; all other element types survive round-trip. |
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
| Tables | cell merging, borders | 🔵 | Round-trip preserved via raw XML capture/re-emission; API-level creation not modeled |
| Sections | page size, margins, orientation | ✅ | |
| Sections | headers and footers | ✅ | Round-trip verified. Rich content (images, formatting) preserved via `_preservedEntries` when not modified via API. |
| Sections | columns | ✅ | `setColumns()` API, round-trip verified |
| Links | hyperlinks (external URLs) | ✅ | |
| Images | inline images with rotation | ✅ | |
| Images | floating (anchored) images | 🔵 | `<wp:anchor>` elements preserved via raw XML capture/re-emission |
| Images | text boxes (w:txbxContent) | ✅ | Text extraction from inline and anchored drawing textboxes is supported |
| Annotations | comments | ✅ | Existing comments round-trip via preservation; minimal `XWPFComment` lookup/text APIs, comment metadata/text editing, and paragraph range comment creation are modeled |
| Annotations | footnotes, endnotes | 🔵 | `word/footnotes.xml` / `word/endnotes.xml` round-trip preserved |
| Fields | TOC, page numbers, mail merge | ✅ | Write/read/round-trip |
| SDT | content controls (block-level and inline) | 🔵 | Block-level `w:sdt` in `w:body` and inline `w:sdt` inside `w:p` preserved via raw XML capture/re-emission. |
| Styles | paragraph style reference (pStyle) | ✅ | `setStyle()`/`getStyleID()` API, round-trip verified. `word/styles.xml` 🔵 preserved + default styles auto-generated. Character/table styles ❌ |
| Track Changes | revision marks | 🔵 | Tracked-change XML (`w:ins`, `w:del`, moves, etc.) is preserved in body/paragraph child order during round-trip; accept/reject/create/edit APIs are not modeled |
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
| Shapes | grouping, connectors, lines | 🔵 | Unknown `p:spTree` children (grpSp, cxnSp, etc.) preserved verbatim via raw XML capture/re-emission |
| Shapes | SmartArt, charts | 🔵 | Preserved as unknown parts |
| Media | video/audio embedding | 🔵 | Non-image `ppt/media/*` round-trip via `_preservedEntries` |
| Animation | animations, transitions | 🔵 | Preserved as unknown parts |
| Theme | layout, master, theme | 🔵 | Not editable, preserved on round-trip |
| Other | pptm macro preservation | ✅ | VBA bytes preserved |

## xls / HSSF (~35%)

| Category | Feature | Status | Notes |
|---|---|---|---|
| Cell values | string, numeric, boolean, blank, error | ✅ | BIFF8 LabelSST/Number/BoolErr/Blank round-trip covered |
| Sheets | multiple sheets, sparse rows/cells, high column indexes | ✅ | |
| Styles | fonts, data formats, alignment, wrap, borders, fills | ⚠️ | Common HSSFFont/HSSFCellStyle round-trip works; not full BIFF style parity |
| Layout | column width, row height, hidden rows/columns, merged regions, freeze panes | ✅ | |
| Formulas | formula text + cached value read | ⚠️ | Existing POI formula fixtures can be read; new BIFF formula token writing and evaluation are not implemented |
| Compatibility | representative Apache POI `.xls` fixtures load | ✅ | Includes basic, styles, formulas, hyperlinks, comments, drawings, images, macro files as load/preservation cases |
| Interop | Java POI bidirectional fixtures | ⚠️ | basic/styles/layout/unicode/comprehensive coverage |
| Preservation | non-Workbook OLE streams, VBA streams, unknown BIFF records | ✅ | Light edits preserve unmodeled streams/records where possible |
| Not modeled | images/shapes/charts/comments/hyperlink editing/filters/pivots | ❌ | Some are load/preservation fixtures, but not public usermodel creation/edit APIs |

## doc / HWPF (~25%)

| Category | Feature | Status | Notes |
|---|---|---|---|
| Reading | OLE2 `.doc` open, FIB/table stream parsing | ✅ | `WordDocument` + `0Table`/`1Table` selection and fallback covered |
| Text | main body text extraction | ✅ | CLX/piece table based extraction with compressed and Unicode text pieces |
| UserModel | Range, Paragraph, CharacterRun | ⚠️ | Paragraph/run splitting and text composition covered for representative fixtures |
| Formatting | character and paragraph properties | ⚠️ | CHPX-derived font name/size/bold/italic/underline/strike plus minimal PAPX fields |
| Extraction | **header/footer and table structures** | ✅ | `getHeaderStoryRange()`, table row/cell iteration implemented |
| Editing | no-op write, append paragraph, simple text replacement | ⚠️ | Limited main-body edit path, not a full Word binary editing engine |
| Preservation | OLE streams/storages, embedded OLE | ✅ | Unedited stream/storage content is preserved in representative fixtures |
| Interop | Java POI bidirectional testing | ⚠️ | Java POI correctly extracts tables and header/footer text from dotnet-poi saved files |
| Not modeled | images/footnotes/comments/fields API | ❌ | Existing streams may be preserved, but these are not usermodel creation/edit features |

## ppt / HSLF (~12%)

| Category | Feature | Status | Notes |
|---|---|---|---|
| Reading | OLE2 `.ppt` open and stream inventory | ✅ | Detects `PowerPoint Document`, `Current User`, and summary streams |
| Records | record tree scan with raw bytes | ✅ | Container/atom hierarchy is retained for preservation |
| Slides | slide count and order | ✅ | Persist-pointer based order covered by representative fixtures |
| Text | TextChars/TextBytes extraction | ✅ | UTF-16LE and CP1252 text atoms supported |
| Editing | no-op write / round-trip | ✅ | OLE2 compound file is preserved through write |
| Preservation | OLE streams/storages, images, comments, OLE | ✅ | Representative preservation fixtures covered |
| Interop | Java POI direction B | ⚠️ | C# fixture generation exists; Java-side assertion still pending |
| Not modeled | slide creation, shape editing, images, animations | ❌ | New presentation authoring is not implemented for HSLF |
