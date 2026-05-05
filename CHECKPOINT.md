# CHECKPOINT

## 2026-05-05 08:xx JST - Phase 7 assessment, formula evaluator dropped, xlsm interop

- Current task: write Phase 7 current-state assessment into AGENTS.md, drop formula evaluator, implement xlsm interop.
- Scope boundary: AGENTS.md updated (English + Japanese), no commit.

### AGENTS.md updates

- Phase 5 (英日両方): "永久凍結" — `XSSFFormulaEvaluator` は既存テスト用に残すが拡張しない。数式評価はスコープ外。
- Phase 7 step 1〜5: 現在地を `[x]`/`[~]`/`[ ]` で明示。モデル層の残差異（fileVersion等）をノートとして記録。
- Phase 7 進捗表（step別パーセンテージ）を追加。

### xlsm interop 実装

C# 側: `WriteForPoiTests.Write_XlsmWithCellsAndVba_CreatesFixtureForPoi`
- `example.xlsm` から VBA バイトを取得（csproj に content item 追加）
- `XSSFWorkbook` で "MacroSheet" シート + A1="from dotnet-poi xlsm" + B1=99.5 + setVBAProject
- `from-dotnet-poi/phase-xlsm-interop.xlsm` に書き出し

Java 側: `ReadFromDotnetTest.readPhaseXlsmInterop`
- `isMacroEnabled()` == true
- Sheet "MacroSheet" + A1 = "from dotnet-poi xlsm" + B1 = 99.5
- OPC パッケージに vbaProject.bin (application/vnd.ms-office.vbaProject) が存在し非空
- `assertNotNull` import 追加

### Verification

- `dotnet test tests/DotnetPoi.Interop.Tests/cs/...` passed (29 tests, was 28).
- `mvn test -Dtest=ReadFromDotnetTest` passed (14 tests, was 13).
- 既存全テストスイート異常なし。

## 2026-05-05 07:xx JST - Namespace, attribute order, and semantic XSSF tests (items 7-12)

- Current task: implement `XMLBEANS_XML_OUTPUT_TODO.md` Implementation Order items 7-12.
- Scope boundary from user: follow AGENTS.md, update `CHECKPOINT.md` while working, and do not commit.

### Items 7 & 8 — Namespace tests and implementation

New file: `tests/DotnetPoi.SS.Tests/Xml/PoiXmlWriterNamespaceTests.cs` (8 tests)

Tests added:
- Default namespace declaration (`xmlns="..."`) on root element
- `xmlns:r` relationship prefix — both the `WriteAttributeString("xmlns:r", ...)` and prefix overload
- Full spreadsheet workbook root pattern: default ns + `xmlns:r`
- Drawing root pattern: `xmlns:xdr`, `xmlns:a`, `xmlns:r` (in POI order)
- No synthetic `main:` prefix: elements written without prefix don't acquire `main:`
- Prefixed elements use caller-supplied prefix, not a synthetic one
- No duplicate namespace declarations when caller writes once

Implementation (item 8): **no production code changes needed**. `PoiXmlWriter` already passes through namespace declarations as plain attributes and does not sort, hoist, or deduplicate them. The tests serve as the specification.

### Item 9 — Attribute order tests

New file: `tests/DotnetPoi.SS.Tests/Xml/PoiXmlWriterAttributeOrderTests.cs` (5 tests)

Tests added:
- Page margins in POI order (`left`, `right`, `top`, `bottom`, `header`, `footer`) — not alphabetical
- Reverse-alphabetical order (`z`, `a`, `m`) — proves no sorting
- `.rels` relationship `Id`, `Type`, `Target` order
- Two sibling elements each with independent attribute order
- Namespace declarations also follow caller order (xdr before a before r)

Implementation: **no production code changes needed**. The writer already preserves caller attribute order.

### Items 10-12 — Semantic XSSF tests using POI integration fixtures

New file: `tests/DotnetPoi.Interop.Tests/cs/PoiIntegrationFixtureTests.cs` (11 tests)

Tests read from `tests/DotnetPoi.Interop.Tests/fixtures/poi-integration/_workbooks/`:
- `poi-integration-shared-strings-basic.xlsx`: 3 sheets (Sheet1/rich test/Sheet3), A1="Lorem", B1=111.0, A2="ipsum", B2=222.0
- `poi-integration-shared-strings-escaping.xlsx`: first cell contains literal `<` (decoded from `&lt;`)
- `poi-integration-styles-formatting.xlsx`: 3 sheets, Sheet1 A1 = "Dates, all 24th November 2006"
- `poi-integration-comments-write-read.xlsx`: 3 sheets (Sheet1/Sheet2/Sheet3, "AllANumbers"/"AllBStrings" are defined names not sheets), A1="A1", B1="B1", A2=22.3, A3=24.5
- `poi-integration-xlsm-vba-preserve.xlsm`: HasMacros=true, 3 sheets (SheetA/SheetB/SheetC), VBA bytes preserved byte-for-byte on round-trip

Fixture path helper: `GetPoiIntegrationFixturePath` traverses up from `AppContext.BaseDirectory` to find `poi-integration/_workbooks/`. If fixture doesn't exist, test message directs user to run Maven generator.

Item 11: no lexical mismatches found; all semantic tests passed without needing additional `PoiXmlWriter` slices.

Item 12: no fixture-specific XML payloads introduced into `XSSFWorkbook`.

### Verification

- `dotnet test tests/DotnetPoi.SS.Tests/...` passed (91 tests, was 78 before items 7/9).
- `dotnet test tests/DotnetPoi.XSSF.Tests/...` passed (26 tests, unchanged).
- `dotnet test tests/DotnetPoi.XWPF.Tests/...` passed (18 tests, unchanged).
- `dotnet test tests/DotnetPoi.XSLF.Tests/...` passed (25 tests, unchanged).
- `dotnet test tests/DotnetPoi.Interop.Tests/cs/...` passed (28 tests, was 17 before items 10-12).
- All commands still show the existing NU1603 warning for `Microsoft.NET.Test.Sdk`.

## 2026-05-05 06:xx JST - Escaping tests and implementation (items 5 & 6)

- Current task: implement `XMLBEANS_XML_OUTPUT_TODO.md` Implementation Order items 5 and 6.
  - Item 5: Add escaping tests for text and attributes (all chars listed in the TODO).
  - Item 6: Implement only the escaping differences proven to diverge from `System.Xml.XmlWriter`.
- Scope boundary from user: follow AGENTS.md, update `CHECKPOINT.md` while working, and do not commit.

### Evidence used

| Context | Char | POI output | Source |
|---|---|---|---|
| Text | `&` | `&amp;` | `xmlbeans-shared-strings-escaping__poi-options.xml` |
| Text | `<` | `&lt;` | same |
| Text | `>` | literal `>` | same — "A&amp;B &lt;C> \"quoted\" 'single'" |
| Text | `"` | literal `"` | same |
| Text | `'` | literal `'` | same |
| Attribute | `&` | `&amp;` | `poi-integration-hyperlinks__xl__worksheets___rels__sheet1.xml.rels` |
| Attribute | `"` | `&quot;` | `poi-integration-styles-formatting__xl__styles.xml` (formatCode) |
| Attribute | `\` | literal `\` | same (formatCode yyyy\\-mm\\-dd) |

`System.Xml.XmlWriter` escaping (measured with a small C# program):
- Text: `>` → `&gt;`, `"` → literal, `'` → literal
- Attributes: `>` → `&gt;`, `"` → `&quot;`, `'` → literal, `\` → literal

### Proven divergences (POI ≠ SXW)

1. **`>` in text content**: POI = literal, SXW = `&gt;`.

### Bugs in original PoiXmlWriter (diverged from both POI and SXW)

2. **`'` in attribute values**: original code produced `&apos;`; both POI and SXW leave `'` literal in double-quoted attributes.

### Implementation (item 6)

Changed `EscapeCore` in `PoiXmlWriter`:
- Removed `case '>'` entirely — `>` is now literal in both text and attributes.
  - Text: matches POI (proven divergence from SXW fixed).
  - Attributes: consistent with XML spec (double-quoted attributes don't require `>` escaping); no POI fixture contradicts this.
- Removed `case '\'' when forAttribute:` — `'` is now literal in attributes (bug fix; matches both POI and SXW).

### Tests added (item 5)

New file: `tests/DotnetPoi.SS.Tests/Xml/PoiXmlWriterEscapingTests.cs`

Text: `&`, `<`, `>` (literal), `"` (literal), `'` (literal), tab, newline, mixed XMLBeans observation.
Attributes: `&`, `"`, `'` (literal), `\`, relationship URL with `&`, format code with `"`, `<`.

3 tests failed before the fix (`GreaterThanInText`, `ApostropheInAttribute`, `MixedSpecialChars`), all pass after.

### Verification

- `dotnet test tests/DotnetPoi.SS.Tests/...` passed (78 tests).
- `dotnet test tests/DotnetPoi.XSSF.Tests/...` passed (26 tests).
- `dotnet test tests/DotnetPoi.XWPF.Tests/...` passed (18 tests).
- `dotnet test tests/DotnetPoi.XSLF.Tests/...` passed (25 tests).
- All commands still show the existing NU1603 warning for `Microsoft.NET.Test.Sdk`.

## 2026-05-05 05:xx JST - Empty element serialization tests and implementation (items 3 & 4)

- Current task: implement the `XMLBEANS_XML_OUTPUT_TODO.md` Implementation Order items 3 and 4.
  - Item 3: Add focused failing tests for empty element serialization.
  - Item 4: Implement empty-element behavior in `PoiXmlWriter`; use stream/text interception if needed.
- Scope boundary from user: follow AGENTS.md, update `CHECKPOINT.md` while working, and do not commit.
- Implementation decision:
  - `PoiXmlWriter` is already a custom text-writer-based implementation (not a wrapper around `System.Xml.XmlWriter`).
    It writes `/>` directly in `WriteEndElement()` when the start tag has not yet been closed, producing `<tag/>` with no space before the slash.
  - The "narrow stream/text interception layer" option from the TODO is satisfied by design: the writer uses `TextWriter` directly rather than delegating to `System.Xml.XmlWriter`.
  - No new production code was needed; the implementation was already correct.
- Completed (item 3, initial pass):
  - Added `PoiXmlWriterEmptyElementTests` with root, nested, prefixed, and single-attributed empty-element cases.
- Completed (item 4, strengthened coverage):
  - Extended `PoiXmlWriterEmptyElementTests` with three additional cases drawn from real OOXML patterns:
    - Multi-attributed empty element: `<Relationship Id="..." Type="..." Target="..."/>` (covers `*.rels` patterns)
    - Prefixed + attributed empty element: `<a:picLocks noChangeAspect="1"/>` (covers drawing namespace patterns)
    - Empty string write before `WriteEndElement`: confirms `WriteString("")` does not prevent the `<tag/>` form.
- Verification:
  - `dotnet test tests/DotnetPoi.SS.Tests/DotnetPoi.SS.Tests.csproj --filter PoiXmlWriterEmptyElementTests` passed (7 tests).
  - `dotnet test tests/DotnetPoi.SS.Tests/DotnetPoi.SS.Tests.csproj` passed (64 tests).
  - `dotnet test tests/DotnetPoi.XSSF.Tests/DotnetPoi.XSSF.Tests.csproj` passed (26 tests).
  - All test commands still show the existing NU1603 package-resolution warning for `Microsoft.NET.Test.Sdk 17.8.2` resolving to `17.9.0`.

## 2026-05-05 04:xx JST - XML writer factory/profile layer

- Current task: implement `XMLBEANS_XML_OUTPUT_TODO.md` Implementation Order 2.
- Scope boundary from user: follow AGENTS.md, keep `CHECKPOINT.md` updated, and do not commit.
- Planned implementation:
  - Add a small factory/profile layer around `PoiXmlWriter` so callers choose XMLBeans spreadsheet-part vs OPC package-part output deliberately.
  - Keep declaration serialization in `PoiXmlWriter`, but avoid hard-coding a single global declaration rule across all OOXML parts.
  - Update XSSF/XWPF/XSLF package writers to create writers through the profile layer instead of directly constructing `PoiXmlWriter`.
  - Add focused xUnit coverage for the factory/profile selection.
- Completed:
  - Added `PoiXmlWriterFactory` with explicit profile creation and OOXML package part classification.
  - Classified `[Content_Types].xml`, `*.rels`, and `docProps/core.xml` as OPC package parts; other XML package entries use the XMLBeans profile.
  - Updated XSSF, XWPF, and XSLF `WriteEntry` helpers to create profiled writers and removed duplicated per-part declaration calls from the writer methods.
  - Left Agile encryption XML on direct `PoiXmlWriter` construction because that XML payload intentionally has no OOXML ZIP part declaration.
- Verification:
  - `dotnet test tests/DotnetPoi.SS.Tests/DotnetPoi.SS.Tests.csproj --filter "PoiXmlWriterFactoryTests|PoiXmlWriterDeclarationProfileTests"` passed.
  - `dotnet test tests/DotnetPoi.SS.Tests/DotnetPoi.SS.Tests.csproj` passed.
  - `dotnet test tests/DotnetPoi.XSSF.Tests/DotnetPoi.XSSF.Tests.csproj` passed.
  - `dotnet test tests/DotnetPoi.XWPF.Tests/DotnetPoi.XWPF.Tests.csproj` passed.
  - `dotnet test tests/DotnetPoi.XSLF.Tests/DotnetPoi.XSLF.Tests.csproj` passed.
  - All dotnet test commands still show the existing NU1603 package-resolution warning for `Microsoft.NET.Test.Sdk 17.8.2` resolving to `17.9.0`.

## 2026-05-05 03:xx JST - XML declaration profile focused tests

- Current task: implement `XMLBEANS_XML_OUTPUT_TODO.md` Implementation Order 1 by adding focused `PoiXmlWriter` tests for XML declaration profiles.
- Scope boundary from user: keep production implementation minimal for now; do not make fixture-specific `XSSFWorkbook` changes; do not commit.
- Target profiles:
  - XMLBeans spreadsheet parts: `<?xml version="1.0" encoding="UTF-8"?>` followed by a newline, no `standalone`.
  - OPC package parts: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` immediately followed by the root element, no forced newline.
- Implementation approach: add a narrow declaration-profile API on `PoiXmlWriter` only, then characterize both profiles with byte-level tests.
- Completed: added `PoiXmlDeclarationProfile` and focused tests in `PoiXmlWriterDeclarationProfileTests`.
- Verification: `dotnet test tests/DotnetPoi.SS.Tests/DotnetPoi.SS.Tests.csproj --filter PoiXmlWriterDeclarationProfileTests` passed; full `dotnet test tests/DotnetPoi.SS.Tests/DotnetPoi.SS.Tests.csproj` passed with the existing NU1603 package-resolution warning.

## 2026-05-05 02:xx JST - Agreed recovery plan for XML parity work

- New direction agreed with user:
  1. Remove the `31e9006 parity` work from production/test code for now, including the fixture-specific XSSF writer changes and associated parity tests/fixtures that forced those changes.
  2. Re-study Java POI XML output behavior and XMLBeans behavior at the correct layer. Add reference code/fixture generators where useful, and write a dedicated Markdown TODO/design file describing observed XMLBeans/POI output patterns, open questions, and the implementation order.
  3. Re-implement the behavior incrementally in `PoiXmlWriter`, with focused failing tests per low-level XML divergence. Keep higher-level `XSSFWorkbook` output POI-model-driven, and preserve unknown/original package parts byte-for-byte where the model does not yet support them.
- Important boundary: XML lexical quirks such as declaration format, empty element form, escaping, attribute order, namespace placement, and whitespace belong in `PoiXmlWriter` or focused helpers. Specific workbook content such as defined names, Office revision GUIDs, local absolute paths, workbook extLst contents, and fixture-specific relationship ordering must not be generalized into `XSSFWorkbook` unless directly backed by the POI model/source behavior.
- Do not commit via LLM.
- Step 1 status: completed in working tree by restoring `src/` and the parity-related `tests/` paths back to `7a4b778` (the parent of `31e9006`). This removes the fixture-specific XSSF writer changes, the added XSSF/XWPF/XSLF parity tests, the expanded Java parity fixture generator, and the extra generated xml-parity fixtures. `dotnet test` passes after the removal.
- Step 2 direction: do not add XMLBeans as a submodule yet. Add Java probe/fixture generator coverage instead and document the work in `XMLBEANS_XML_OUTPUT_TODO.md`. Use XMLBeans 5.3.0 source jars/tagged source only if executable probes are not enough.
- Step 3 status: added `XMLBEANS_XML_OUTPUT_TODO.md`, `XmlBeansOutputProbeTest.java`, and initial generated XMLBeans probe fixtures under `tests/DotnetPoi.Interop.Tests/fixtures/xmlbeans-output/`. Verified with `mvn test -f tests/DotnetPoi.Interop.Tests/java/pom.xml -Dtest=XmlBeansOutputProbeTest`.
- Additional fixture strategy: use upstream POI integration-level tests as scenario sources to improve fixture realism. Added `POI_INTEGRATION_FIXTURE_TODO.md` and updated `XMLBEANS_XML_OUTPUT_TODO.md` with a step-by-step plan. Keep POI-derived fixtures semantic-first; do not turn fixture-specific XML into generalized `XSSFWorkbook` output.
- Fixture collection progress: surveyed POI XSSF/model/openxml4j tests and `poi/test-data`. Shortlisted 8 selected candidates in `POI_INTEGRATION_FIXTURE_TODO.md`: shared strings basic, shared strings escaping, styles formatting, comments write/read, pictures multi-sheet, relationships hyperlinks/comments, xlsm VBA preserve, and rich text whitespace preserve. First implementation target is `poi-integration-shared-strings-basic`.
- Fixture collection progress: added `PoiIntegrationFixtureGeneratorTest.java` and generated the first POI-derived integration fixture, `poi-integration-shared-strings-basic`, from POI `sample.xlsx` / `TestSharedStringsTable.testReadWrite` scenario. Output is one xlsx package plus 16 extracted XML/rels files under `tests/DotnetPoi.Interop.Tests/fixtures/poi-integration/`. Verified with `mvn test -f tests/DotnetPoi.Interop.Tests/java/pom.xml -Dtest=PoiIntegrationFixtureGeneratorTest`.
- Fixture collection progress: extended `PoiIntegrationFixtureGeneratorTest.java` with `poi-integration-shared-strings-escaping` from POI `TestSharedStringsTable.testBug48936` and `poi-integration-styles-formatting` from POI `TestStylesTable.testLoadSaveLoad`. Total generated POI integration fixture files are now 40: three workbook packages plus extracted XML/rels. Verified with the same Maven command.
- Fixture collection progress: extended `PoiIntegrationFixtureGeneratorTest.java` with `poi-integration-comments-write-read`, `poi-integration-pictures-multi-sheet`, and `poi-integration-relationships-hyperlinks-comments`. Total generated POI integration fixture files are now 90: six workbook packages plus extracted XML/rels. Verified with `mvn test -f tests/DotnetPoi.Interop.Tests/java/pom.xml -Dtest=PoiIntegrationFixtureGeneratorTest`.
- Fixture collection progress: extended `PoiIntegrationFixtureGeneratorTest.java` with `poi-integration-xlsm-vba-preserve` from POI `TestXSSFBugs.bug45431` and `poi-integration-rich-text-space-preserve` from POI `TestXSSFRichTextString` whitespace cases. Extraction now includes `.bin` and `.vml` as well as XML/rels so macro/VML preservation fixtures include `xl/vbaProject.bin` and `xl/drawings/vmlDrawing1.vml`. Total POI integration fixture files are now 123: eight workbook packages plus extracted package entries. Verified with the same Maven command.
- Fixture collection progress: extended `PoiIntegrationFixtureGeneratorTest.java` with `poi-integration-defined-names-print-titles`, `poi-integration-hyperlinks`, `poi-integration-sheet-layout`, and `poi-integration-formula-recalculation`. These cover defined names/print titles, external hyperlink relationships and query escaping, row/column layout with panes/grouping/merged regions, and workbook formula recalculation XML. Total POI integration fixture files are now 165: twelve workbook packages plus extracted package entries. Verified with `mvn test -f tests/DotnetPoi.Interop.Tests/java/pom.xml -Dtest=PoiIntegrationFixtureGeneratorTest`.
- Fixture analysis progress: updated `XMLBEANS_XML_OUTPUT_TODO.md` with observations from the XMLBeans probes and 12 POI integration fixture cases. Key split: spreadsheet XMLBeans parts use UTF-8 declaration plus newline and no `standalone`, while OPC package-level XML parts such as `[Content_Types].xml`, `.rels`, and core properties use `standalone="yes"` with no forced newline before the root. Recorded `PoiXmlWriter` implementation order: declaration profiles, `<tag/>` empty elements, escaping, namespace behavior, attribute-order preservation, then semantic fixture tests and unknown-part preservation.

## 2026-05-05 02:xx JST - Review of recent XSSF writer/parity work

- User raised concern that recent commits may have made XSSF workbook XML writing ad hoc and destabilizing.
- Reviewed last 5 commits. Main risky commit is `31e9006 parity`; `3830f84` adjusts generated fixtures/timestamps.
- `dotnet test` passes locally, but `XSSFWorkbook.WriteWorkbook` now contains fixture-specific-looking constants: Office revision namespaces, GUIDs, a local absolute `x15ac:absPath`, hard-coded defined names, workbook window values, and extension list data.
- Assessment: the concern is valid. The tests currently prove parity for a narrow fixture, not a general POI-faithful writer. Recommended next step is to quarantine this behavior behind preservation of original package parts or fixture-only tests, then revert generalized writer output to minimal/POI-derived data.

## 2026-05-05 01:54 JST - XML parity CI drift

- GitHub Actions `Verify XML Parity Fixtures` failed because `XmlParityFixtureGeneratorTest` rewrites `xlsm-basic` by opening `tests/test-files/example.xlsm` with Apache POI and saving it through `XSSFWorkbook.write()`.
- The committed `xlsm-basic__*.xml` fixtures are intentionally hybrid: DotnetPoi regenerates workbook/content-types/relationships but preserves unchanged macro workbook parts such as doc props, styles, shared strings, worksheets, drawings, theme, and calcChain.
- Fix: updated `generateMacroEnabledXlsm` to first generate the POI package, then overlay the preserved xlsm entries from the source workbook so CI regenerates the same hybrid fixture set.
- Verification: `dotnet test tests/DotnetPoi.XSSF.Tests/DotnetPoi.XSSF.Tests.csproj --filter XmlParity_XlsmBasic_MatchesPoiFixtures` passes. Local Maven is not installed, so the exact GitHub workflow command still needs CI or a Maven-equipped environment.
