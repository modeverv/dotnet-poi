# Phase -1 Step 3: XmlWriter Divergence Candidates

Source fixtures: `tests/DotnetPoi.Interop.Tests/fixtures/xml-parity`

Scope: all 60 extracted XML relationship/package parts, excluding `_workbooks/*.xlsx`.

This document records the XMLBeans/Apache POI byte-level output patterns that C# code must reproduce. Each item below is a candidate for a failing `PoiXmlWriter` test in Step 4.

## High Priority Candidates

### 1. XML Declaration Shape

POI emits exactly one of two declaration forms:

- Same-line package/core form: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` followed immediately by the root element.
- Spreadsheet/app form: `<?xml version="1.0" encoding="UTF-8"?>` followed by a newline, then the root element.

Candidate risk:

- `XmlWriter` can omit `standalone="yes"`.
- `XmlWriter` can use lowercase or runtime encoding names that differ from `UTF-8`.
- `XmlWriter` can place the root on a different line depending on settings.

Locations:

- `standalone="yes"` and root on same line: all `[Content_Types].xml`, all `*.rels`, and all `docProps/core.xml` files.
- no standalone and root on next line: all `docProps/app.xml`, `xl/workbook.xml`, `xl/sharedStrings.xml`, `xl/styles.xml`, `xl/worksheets/*.xml`, `xl/comments1.xml`, and `xl/drawings/*.xml` files.

Observed counts:

- 26 files use `standalone="yes"` with root on the declaration line.
- 34 files omit `standalone` and put the root on the next line.

### 2. Empty Element Closing

POI emits empty elements as `<tag/>` with no space before `/`.

Candidate risk:

- `System.Xml.XmlWriter` commonly emits `<tag />`.
- `WriteFullEndElement()` emits `<tag></tag>`, which is also not byte-identical.

Observed counts:

- 255 empty elements across 45 files.

Representative locations:

- `[Content_Types].xml`: `Default`, `Override`.
- `*.rels`: `Relationship`.
- `xl/styles.xml`: `numFmts`, `sz`, `color`, `name`, `family`, `scheme`, `patternFill`, `left`, `right`, `top`, `bottom`, `diagonal`, `xf`.
- `xl/workbook.xml`: `workbookPr`, `workbookView`, `sheet`.
- `xl/worksheets/*.xml`: `dimension`, `sheetView`, `sheetFormatPr`, `c`, `row`, `pageMargins`, `drawing`, `legacyDrawing`.
- `xl/sharedStrings.xml`: empty `sst` and empty `t` variants.
- `xl/comments1.xml`: empty `author`.
- `xl/drawings/drawing1.xml`: `xdr:cNvPr`, `a:picLocks`, `a:blip`, `a:fillRect`, `a:off`, `a:ext`, `a:avLst`, `xdr:clientData`.

Step 4 should start here: this is the most universal and easiest byte-level divergence to lock down.

### 3. Attribute Ordering

POI/XMLBeans attribute order is stable and must be preserved by the porting layer. The writer must not sort attributes and must not rely on dictionaries with unstable enumeration order.

Candidate risk:

- Code that emits attributes from unordered data structures will produce valid XML but byte-different output.
- Some POI orders are not the order a .NET implementation might naturally choose.

Required observed orders:

- `Default`: `ContentType`, `Extension`.
- `Override`: `ContentType`, `PartName`.
- `Relationship`: `Id`, `Target`, `Type`.
- `workbook`: `xmlns`, then `xmlns:r`.
- `workbookPr`: `date1904`.
- `workbookView`: `activeTab`.
- `sheet`: `name`, `r:id`, `sheetId`.
- `worksheet`: `xmlns`, then optional `xmlns:r`.
- `dimension`: `ref`.
- `sheetView`: `workbookViewId`, then optional `tabSelected`.
- `sheetFormatPr`: `defaultRowHeight`.
- `c`: `r`, then optional `t`, then `s`.
- `pageMargins`: `bottom`, `footer`, `header`, `left`, `right`, `top`.
- `sst`: `count`, `uniqueCount`, `xmlns`.
- `xf`: `numFmtId`, `fontId`, `fillId`, `borderId`, then optional `xfId`/`applyNumberFormat`.
- `numFmt`: `formatCode`, `numFmtId`.
- `xdr:wsDr`: `xmlns:xdr`, `xmlns:a`, `xmlns:r`.
- `xdr:twoCellAnchor`: `editAs`.
- `xdr:cNvPr`: `id`, `name`, `descr`.
- `a:picLocks`: `noChangeAspect`.
- `a:blip`: `r:embed`.
- `a:off`: `x`, `y`.
- `a:ext`: `cx`, `cy`.
- `a:prstGeom`: `prst`.
- `dcterms:created`: `xsi:type`.

### 4. Explicit Zero and Default Attributes

POI writes several zero/default values explicitly. The .NET port must not treat these as optional defaults.

Candidate risk:

- Object serializers, schema-aware writers, or hand-written convenience helpers may omit `0`, `false`, or default-height attributes.

Locations:

- `sharedStrings.xml`: `sst count="0" uniqueCount="0"` when no shared strings exist.
- `styles.xml`: `numFmts count="0"`.
- `styles.xml`: `xf numFmtId="0" fontId="0" fillId="0" borderId="0"`.
- `styles.xml`: `cellXfs/xf xfId="0"`.
- `workbook.xml`: `workbookPr date1904="false"`.
- `workbook.xml`: `workbookView activeTab="0"`.
- `worksheets/*.xml`: `sheetView workbookViewId="0"`.
- `worksheets/*.xml`: `sheetView tabSelected="true"` on selected sheets only.
- `worksheets/*.xml`: `sheetFormatPr defaultRowHeight="15.0"`.
- `worksheets/*.xml`: cells often include `s="0"`.
- `drawing1.xml`: `xdr:col`, `xdr:colOff`, `xdr:row`, `xdr:rowOff` text values of `0`.
- `drawing1.xml`: `a:off x="0" y="0"`.
- `drawing1.xml`: `a:ext cx="0" cy="0"`.

### 5. Namespace Declaration Placement and Prefix Order

POI places namespace declarations on the root element and uses minimal scoped declarations in these fixtures.

Candidate risk:

- `XmlWriter` may duplicate namespace declarations or hoist/relocate them if namespace APIs are used indirectly.
- Prefix declaration order is byte-visible.

Observed root namespace patterns:

- `[Content_Types].xml`: default `http://schemas.openxmlformats.org/package/2006/content-types`.
- `*.rels`: default `http://schemas.openxmlformats.org/package/2006/relationships`.
- `docProps/core.xml`: `xmlns:cp`, `xmlns:dc`, `xmlns:dcterms`, `xmlns:xsi` in that order.
- `docProps/app.xml`: default `http://schemas.openxmlformats.org/officeDocument/2006/extended-properties`.
- `xl/workbook.xml`: default spreadsheet namespace, then `xmlns:r`.
- normal worksheets: default spreadsheet namespace only.
- `namespaces__xl__worksheets__sheet1.xml`: default spreadsheet namespace, then `xmlns:r`.
- `xl/styles.xml`, `xl/sharedStrings.xml`, `xl/comments1.xml`: default spreadsheet namespace only.
- `xl/drawings/drawing1.xml`: `xmlns:xdr`, `xmlns:a`, `xmlns:r` in that order.

### 6. Element Ordering

Several parts are order-sensitive for byte parity even though semantic readers may accept alternate ordering.

Candidate risk:

- Sorting by part name or relationship type can produce valid OOXML but differ from POI byte output.

Required observed orders:

- Root `.rels`: office document, extended properties, core properties.
- `xl/_rels/workbook.xml.rels`: shared strings, styles, then worksheets in sheet order.
- `worksheets/_rels/sheet1.xml.rels`: drawing, comments, VML drawing.
- `[Content_Types].xml`: `Default` entries first, then `Override` entries; within `namespaces`, defaults are `png`, `rels`, `vml`, `xml`.
- `workbook.xml`: `workbookPr`, `bookViews`, `sheets`.
- `workbook.xml` sheets: sheet order as created; `Alpha`, `Beta`, `Gamma` in the multi-sheet fixture.
- `worksheet.xml`: `dimension`, `sheetViews`, `sheetFormatPr`, `sheetData`, `pageMargins`, then optional `drawing` and `legacyDrawing`.
- `styles.xml`: `numFmts`, `fonts`, `fills`, `borders`, `cellStyleXfs`, `cellXfs`.
- `core.xml`: `dcterms:created`, then `dc:creator`.
- `comments1.xml`: `authors`, then `commentList`.
- `drawing1.xml`: `xdr:from`, `xdr:to`, `xdr:pic`, `xdr:clientData` inside `xdr:twoCellAnchor`.

### 7. Whitespace and Newlines Inside XML Bodies

Most fixtures are unindented single-line XML after the declaration, but there are meaningful exceptions.

Candidate risk:

- Pretty printing or indentation will break byte parity.
- Removing generator-emitted newlines can also break parity.

Observed patterns:

- `standalone="yes"` package parts put declaration and root on the same physical line.
- most non-standalone spreadsheet parts put a newline after the declaration and then use compact XML.
- `inline-strings__xl__worksheets__sheet1.xml` contains newlines around `sheetData`, `row`, and `c` content.
- `docProps/app.xml` contains a newline after the declaration and otherwise compact XML.

### 8. Text and Scalar Formatting

POI's scalar formatting is byte-visible.

Candidate risk:

- Numeric conversion through current culture, invariant culture differences, or canonicalization can change values.

Observed values to preserve exactly:

- zero numeric cells: `0.0`.
- negative zero numeric cell: `-0.0`.
- date serial: `25569.375`.
- default row height: `15.0`.
- booleans as cell values: `1`.
- boolean attributes: lowercase `true`/`false`.
- formulas as raw text: `SUM(1,2,3)`.

## Fixture-Family Coverage

The following family-level candidates cover every extracted fixture file:

- `cell-empty`: empty shared string text, empty row/cell markup, baseline package/workbook/styles relationships.
- `cell-types`: string, numeric, boolean, date serial, and formula cells; custom number format style.
- `cell-zero`: zero and negative-zero numeric cell serialization.
- `inline-strings`: inline string cells and body newlines inside `sheetData`.
- `multi-sheet`: sheet order, workbook relationship order, and multi-sheet content types.
- `namespaces`: prefixed worksheet relationships, comments, drawings, image relationship, VML content type, and DrawingML namespace order.

## Step 4 Test Backlog

Recommended failing-test order:

1. Declaration tests: standalone same-line package part, non-standalone newline spreadsheet part.
2. Empty-element tests: `<tag/>` with no space and no full end tag.
3. Attribute-order tests: `Relationship`, `Override`, `sheet`, `c`, `xf`, `pageMargins`, and drawing elements.
4. Zero/default tests: `count="0"`, `date1904="false"`, `activeTab="0"`, `s="0"`, `workbookViewId="0"`, drawing `x/y/cx/cy="0"`.
5. Namespace-order tests: workbook default + `r`, core props prefixes, drawing `xdr/a/r`.
6. Element-order tests: workbook relationships, content types, worksheet body, styles, comments/drawing.
7. Whitespace tests: same-line package parts, declaration newline parts, inline string body newlines.
8. Scalar-format tests: `0.0`, `-0.0`, `15.0`, lowercase booleans.
