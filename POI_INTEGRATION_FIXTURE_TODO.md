# POI Integration Fixture TODO

This file tracks the plan to improve fixture realism by borrowing scenario shapes from upstream Apache POI integration-level tests.

## Purpose

Use upstream POI tests to choose realistic OOXML scenarios, then recreate those scenarios in dotnet-poi's Java fixture generator. These fixtures should improve semantic compatibility and round-trip coverage without pushing fixture-specific XML payloads into production writers.

## Rules

- Read upstream POI tests before selecting a scenario.
- Record upstream path, test method, scenario summary, and why it matters.
- Recreate scenario shape in `tests/DotnetPoi.Interop.Tests/java`; do not blindly copy POI assertions.
- Prefer semantic C# assertions over byte-level XML assertions.
- Byte-level assertions are allowed only for isolated XMLBeans/PoiXmlWriter lexical behavior.
- Unknown or unsupported package parts should be preserved byte-for-byte rather than reimplemented ad hoc.

## Candidate Table

| Status | Upstream POI path | Method / scenario | Why it matters | Dotnet fixture/test |
|---|---|---|---|---|
| generated | `poi/poi-ooxml/src/test/java/org/apache/poi/xssf/model/TestSharedStringsTable.java` | `testReadWrite` with `sample.xlsx` | Baseline sharedStrings round-trip and package realism | Java fixture case: `poi-integration-shared-strings-basic` |
| generated | `poi/poi-ooxml/src/test/java/org/apache/poi/xssf/model/TestSharedStringsTable.java` | `testBug48936` | Stress shared string escaping / CDATA-sensitive content without depending on high-level workbook XML hacks | Java fixture case: `poi-integration-shared-strings-escaping` |
| generated | `poi/poi-ooxml/src/test/java/org/apache/poi/xssf/model/TestStylesTable.java` | `testLoadSaveLoad` with `Formatting.xlsx` | Styles table, numFmts, fonts, fills, borders, and `styles.xml` round-trip | Java fixture case: `poi-integration-styles-formatting` |
| generated | `poi/poi-ooxml/src/test/java/org/apache/poi/xssf/model/TestCommentsTable.java` | `writeRead` with `WithVariousData.xlsx` | Comments, VML/drawing relationships, multi-sheet comment add/modify round-trip | Java fixture case: `poi-integration-comments-write-read` |
| generated | `poi/poi-ooxml/src/test/java/org/apache/poi/xssf/usermodel/TestXSSFPicture.java` | same image referred to by multiple sheets | Drawing relationships, media reuse, multiple sheets sharing images | Java fixture case: `poi-integration-pictures-multi-sheet` |
| generated | `poi/poi-ooxml/src/test/java/org/apache/poi/openxml4j/opc/TestRelationships.java` | `testFetchFromCollection` with `ExcelWithHyperlinks.xlsx` | Sheet relationship collections, hyperlinks, comments relationship IDs | Java fixture case: `poi-integration-relationships-hyperlinks-comments` |
| generated | `poi/poi-ooxml/src/test/java/org/apache/poi/xssf/usermodel/TestXSSFBugs.java` | bug 45431 macro carry-over | Macro-enabled package preservation and `vbaProject.bin` relationship/content type | Java fixture case: `poi-integration-xlsm-vba-preserve` |
| generated | `poi/poi-ooxml/src/test/java/org/apache/poi/xssf/usermodel/TestXSSFRichTextString.java` | `xml:space="preserve"` cases | Whitespace preservation in shared strings/rich text; likely `PoiXmlWriter` lexical relevance | Java fixture case: `poi-integration-rich-text-space-preserve` |
| generated | `poi/poi-ooxml/src/test/java/org/apache/poi/xssf/usermodel/TestXSSFName.java` / workbook print-title behavior | defined names and print titles | Workbook-level defined names, sheet-scoped names, quoted sheet references, and built-in `_xlnm.Print_Titles` output | Java fixture case: `poi-integration-defined-names-print-titles` |
| generated | `poi/poi-ooxml/src/test/java/org/apache/poi/xssf/usermodel/TestXSSFHyperlink.java` | `testCreate` style URL hyperlink creation | External hyperlink relationships, query escaping, and sheet hyperlink references | Java fixture case: `poi-integration-hyperlinks` |
| generated | `poi/poi-ooxml/src/test/java/org/apache/poi/xssf/usermodel/TestXSSFSheet.java` / sheet layout scenarios | row/column layout, panes, grouping, merged regions | Sheet dimension/layout XML: cols, hidden columns, outline rows, freeze pane, selections, and mergeCells | Java fixture case: `poi-integration-sheet-layout` |
| generated | `poi/poi-ooxml/src/test/java/org/apache/poi/xssf/usermodel/TestXSSFWorkbook.java` / formula recalculation behavior | workbook force formula recalculation | Formula cell XML plus workbook `calcPr` output for recalculation-on-load behavior | Java fixture case: `poi-integration-formula-recalculation` |
| later | `poi/poi-ooxml/src/test/java/org/apache/poi/openxml4j/opc/TestPackageCoreProperties.java` | `testAlternateCorePropertyTimezones` | Core properties timestamp normalization; useful but not XSSF-first | TBD |
| later | `poi/test-data/integration/*.xlsx` | stress workbooks | Real-world package shapes; likely too broad for first pass | TBD |

## Execution Steps

1. [x] Survey upstream POI tests with `rg` and shortlist candidates.
2. [x] Add exact candidate rows to this file.
3. [x] Pick one scenario and build a Java fixture generator case.
4. [x] Generate package and extracted XML fixtures.
5. [ ] Add C# semantic compatibility tests.
6. [ ] Only then decide whether any finding belongs in `PoiXmlWriter`.

## First Implementation Choice

Start with `poi-integration-shared-strings-basic`.

Reasons:

- It is small and already relies on `sample.xlsx`.
- It exercises shared strings and package read/write without drawings, macros, or styles complexity.
- The C# semantic assertions can be simple: read values, verify shared string cells survive, and ensure generated XML parts are present.
- It is unlikely to tempt fixture-specific `XSSFWorkbook` output changes.

Generated outputs:

- `tests/DotnetPoi.Interop.Tests/fixtures/poi-integration/_workbooks/poi-integration-shared-strings-basic.xlsx`
- `tests/DotnetPoi.Interop.Tests/fixtures/poi-integration/poi-integration-shared-strings-basic__*.xml`
- `tests/DotnetPoi.Interop.Tests/fixtures/poi-integration/poi-integration-shared-strings-basic__*.rels`

Verification:

```bash
mvn test -f tests/DotnetPoi.Interop.Tests/java/pom.xml -Dtest=PoiIntegrationFixtureGeneratorTest
```

Notes from first fixture:

- Produces 17 files total: one workbook package plus extracted XML/relationship parts.
- Includes `xl/sharedStrings.xml` with rich text runs and `xml:space="preserve"` examples.
- Includes stable core properties from the source workbook; no current-time timestamp drift observed in `docProps/core.xml`.

Additional generated cases:

- `poi-integration-shared-strings-escaping`
  - Source: POI `48936-strings.txt` recreated into a deterministic new workbook.
  - Core properties are pinned to `2007-01-02T03:04:05Z` to avoid current-time fixture drift.
  - Generated 9 extracted XML/rels files plus one workbook package.
- `poi-integration-styles-formatting`
  - Source: POI `Formatting.xlsx`, matching `TestStylesTable.testLoadSaveLoad`.
  - Captures custom numFmts, fonts, fills, borders, cellXfs, tableStyles, and XML escaping in format codes.
  - Generated 12 extracted XML/rels files plus one workbook package.
- `poi-integration-comments-write-read`
  - Source: POI `WithVariousData.xlsx`, matching `TestCommentsTable.writeRead`.
  - Modifies one existing comment and adds a new comment on another sheet.
  - Captures `comments1.xml`, `comments2.xml`, sheet rels, drawing XML, and multi-sheet comment package structure.
  - Generated 17 extracted XML/rels files plus one workbook package.
- `poi-integration-pictures-multi-sheet`
  - Source: recreated from `TestXSSFPicture` multi-sheet image reuse scenario.
  - Uses deterministic in-memory picture bytes and pinned core properties.
  - Captures two drawing parts, two drawing rels, sheet rels, media reuse relationships, and workbook package structure.
  - Generated 16 extracted XML/rels files plus one workbook package.
- `poi-integration-relationships-hyperlinks-comments`
  - Source: POI openxml4j `ExcelWithHyperlinks.xlsx`, matching `TestRelationships.testFetchFromCollection`.
  - Captures sheet relationships for hyperlinks and comments.
  - Generated 14 extracted XML/rels files plus one workbook package.
- `poi-integration-xlsm-vba-preserve`
  - Source: POI `45431.xlsm`, matching `TestXSSFBugs.bug45431`.
  - Captures macro-enabled content types, workbook relationships, `xl/vbaProject.bin`, and `xl/drawings/vmlDrawing1.vml`.
  - Generated 15 extracted package entries plus one workbook package.
- `poi-integration-rich-text-space-preserve`
  - Source: recreated from `TestXSSFRichTextString` whitespace preservation cases.
  - Captures `xml:space="preserve"` on leading/trailing spaces, tabs, trailing newlines, and multi-run rich text.
  - Generated 9 extracted XML/rels files plus one workbook package.
- `poi-integration-defined-names-print-titles`
  - Source: recreated from POI XSSF defined-name and print-title behavior.
  - Captures workbook-level defined names, sheet-scoped built-in print-title names, quoted sheet names, and workbook relationship structure.
  - Generated 10 extracted XML/rels files plus one workbook package.
- `poi-integration-hyperlinks`
  - Source: recreated from `TestXSSFHyperlink.testCreate` style URL hyperlink creation.
  - Captures five external hyperlink relationships and XML escaping for URL query parameters.
  - Generated 10 extracted XML/rels files plus one workbook package.
- `poi-integration-sheet-layout`
  - Source: recreated from POI XSSF sheet layout behaviors.
  - Captures row heights, column widths, hidden columns, grouped rows, freeze panes, selections, and merged cells.
  - Generated 9 extracted XML/rels files plus one workbook package.
- `poi-integration-formula-recalculation`
  - Source: recreated from POI workbook formula recalculation behavior.
  - Captures formula cell XML and workbook `calcPr fullCalcOnLoad="true"` output.
  - Generated 9 extracted XML/rels files plus one workbook package.

Current generated fixture total:

- 12 workbook packages under `tests/DotnetPoi.Interop.Tests/fixtures/poi-integration/_workbooks/`
- 153 extracted package entries under `tests/DotnetPoi.Interop.Tests/fixtures/poi-integration/`
- 165 files total

Latest verification:

```bash
mvn test -f tests/DotnetPoi.Interop.Tests/java/pom.xml -Dtest=PoiIntegrationFixtureGeneratorTest
```

## Notes

- This complements `XMLBEANS_XML_OUTPUT_TODO.md`.
- This work should not resurrect the previous fixture-specific `XSSFWorkbook.WriteWorkbook` behavior.
