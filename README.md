# dotnet-poi

An **unofficial**, faithful port of [Apache POI](https://poi.apache.org/) for .NET.

[![CI](https://github.com/modeverv/dotnet-poi/actions/workflows/ci.yml/badge.svg)](https://github.com/modeverv/dotnet-poi/actions/workflows/ci.yml)
[![Examples](https://github.com/modeverv/dotnet-poi/actions/workflows/examples.yml/badge.svg)](https://github.com/modeverv/dotnet-poi/actions/workflows/examples.yml)
[![XML Parity](https://github.com/modeverv/dotnet-poi/actions/workflows/xml-parity-fixtures.yml/badge.svg)](https://github.com/modeverv/dotnet-poi/actions/workflows/xml-parity-fixtures.yml)
![License](https://img.shields.io/badge/license-Apache%202.0-blue)
![Status](https://img.shields.io/badge/status-WIP-yellow)
![Phase](https://img.shields.io/badge/phase-4%20HSSF%20xls%20bootstrap%20%E2%80%94%20in%20progress-yellow)

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

### Current Phase: Phase 4 — HSSF xls bootstrap in progress

| Phase | Description | Target | Status |
|---|---|---|---|
| **-1** | **XML output parity (Java vs .NET)** | **—** | ✅ Done |
| **0** | **xlsx write (string / number)** | **v0.1** | ✅ Done |
| **1** | **xlsx read (string / number)** | **v0.2** | ✅ Done |
| **2** | **Styles & formatting (font, color, border)** | **v0.3** | ✅ Done |
| **2.5** | **Images & drawing (XSSFPicture, XSSFDrawing)** | **v0.35** | ✅ Done |
| **3** | **SS common interface (IWorkbook / ISheet / IRow / ICell)** | **v0.4** | ✅ Done |
| **3.1** | **Image rotation (XSSFPicture.setRotation/getRotation)** | **—** | ✅ Done |
| **3.2** | **docx support (XWPF — text, images, rotation)** | **—** | ✅ Done |
| **3.3** | **pptx support (XSLF — slides, images, rotation, flip)** | **—** | ✅ Done |
| **3.4** | **AGILE encryption (XSSF write/read decrypt path)** | **—** | ✅ Done |
| **7** | **Gleaning — cell type coverage (Boolean, Formula cached, Error read/write)** | **—** | ✅ Done |
| **5.1** | **Formula save/read (`<f>` round-trip)** | **—** | ✅ Done |
| **5.2** | **Force formula recalculation (`calcPr fullCalcOnLoad`)** | **—** | ✅ Done |
| **5.3** | **Representative formula evaluation functions** | **—** | ✅ Done |
| 4 | POIFS + HSSF (xls read/write) | v0.5 | 🚧 HSSF bootstrap |
| 5 | Full FormulaEvaluator parity | v1.0 | 🚧 Partial |
| **6** | **HWPF (.doc text read) + HSLF (.ppt slide/text read)** | **v1.x** | ✅ Done (read-only MVP) |

Note: Formula evaluation and setting formula result values are intentionally omitted in the library for now; formulas are preserved as text and cached results are only handled when present.

Minimum bar if POIFS is considered “full” (to unblock HWPF/HSLF work):

- Read/write OLE2 header, FAT, mini FAT, and DIFAT chains for multi-stream files
- Directory tree with storage/stream entries, sibling ordering, timestamps, and CLSIDs
- Mini stream support with cutoff behavior and mini stream allocation
- Sector allocation/chain validation for non-contiguous and large streams
- Stream APIs for random access read/write (seek, length, overwrite)
- CodePage handling and UTF-16LE name storage rules
- Graceful handling of unknown streams/properties without data loss
- Test fixtures:
  - round-trip multiple streams with mixed sizes (mini + regular)
  - verify directory tree ordering matches POI’s comparator
  - Java POI interop: POI reads dotnet-poi CFB, dotnet-poi reads POI CFB

### Phase 0 — Class Progress

| Class | Ported | Tested | Notes |
|---|---|---|---|
| `XSSFWorkbook` | ✅ | ✅ | Minimal `.xlsx` package write |
| `XSSFSheet` | ✅ | ✅ | Minimal sheet creation and row access |
| `XSSFRow` | ✅ | ✅ | Minimal row creation and cell access |
| `XSSFCell` | ✅ | ✅ | String, numeric, boolean, error, formulas with cached values |
| `XSSFCreationHelper` | ✅ | ✅ | Data format, anchors, formula evaluator |

Legend: ✅ Done / 🚧 In Progress / ⬜ Not started

### Byte-Level Parity Scope

Current byte-level XML parity is guaranteed at the `PoiXmlWriter` fixture layer, using Apache POI/XMLBeans-generated XSSF `.xlsx` XML parts under `tests/DotnetPoi.Interop.Tests/fixtures/xml-parity/`.

This does **not** currently mean whole-file byte parity for `.doc`, `.docx`, `.xls`, `.xlsx`, `.ppt`, `.pptx`, `.docm`, `.xlsm`, or `.pptm`:

- `.xls`, `.doc`, and `.ppt` are OLE2/binary formats, not XML package formats.
- `.xlsx` XML writer behavior is fixture-gated; full ZIP byte parity is not claimed because ZIP metadata and timestamps vary.
- `.docx`, `.pptx`, `.docm`, `.xlsm`, and `.pptm` currently have semantic round-trip and Apache POI interop coverage where implemented, but do not yet have format-specific XML fixture byte-parity suites comparable to the XSSF `xml-parity/` fixtures.
- Agile encryption keeps the normal OOXML parts on the same XML writer foundation before encryption, but encrypted OLE2 output is intentionally not byte-for-byte comparable with Apache POI because salts, keys, encrypted payload bytes, sector layout, and package metadata vary. The Agile `EncryptionInfo` XML is generated in a known Excel/POI-compatible shape, not asserted as POI byte-identical.

### Phase -1 Foundation

Phase -1 is complete. The project now has a `PoiXmlWriter` foundation for reproducing Apache POI/XMLBeans OOXML output at byte-level fidelity.

What is locked down:

- Java Apache POI fixture generation under `tests/DotnetPoi.Interop.Tests/fixtures/xml-parity/`
- byte-level fixture comparisons for XML declaration shape, empty element style, attribute order, namespace order, explicit zero/default attributes, element order, whitespace, and scalar formatting
- a source gate test that fails if production code bypasses `PoiXmlWriter` with direct XML APIs such as `XmlWriter`, `XDocument`, `XElement`, `XmlDocument`, or `XmlSerializer`

For future work, the XML parity tests must stay green:

```bash
dotnet test tests/DotnetPoi.SS.Tests/DotnetPoi.SS.Tests.csproj --filter PoiXmlWriter
```

### Phase 0 Verification

Phase 0 is complete for the first writable surface: create a workbook, create sheets/rows/cells, write string and numeric values, and save an `.xlsx` file.

Verification currently covers:

- unit tests for the Phase 0 XSSF API and generated OOXML parts
- XML parity tests for the low-level `PoiXmlWriter` fixtures captured from Apache POI/XMLBeans
- Java interop in the write direction: dotnet-poi writes an `.xlsx`, then Apache POI reads and asserts the cell values
- a runnable example under `examples/Phase0WriteExample`

Commands:

```bash
dotnet test tests/DotnetPoi.XSSF.Tests/DotnetPoi.XSSF.Tests.csproj --no-restore
dotnet test tests/DotnetPoi.SS.Tests/DotnetPoi.SS.Tests.csproj --no-restore
dotnet test tests/DotnetPoi.Interop.Tests/cs/DotnetPoi.Interop.Tests.csproj --no-restore --filter Category=WriteForPoi
mvn test -f tests/DotnetPoi.Interop.Tests/java/pom.xml -Dtest=ReadFromDotnetTest
dotnet run --project examples/Phase0WriteExample/Phase0WriteExample.csproj
```

The example writes:

```text
examples/output/phase0-write-example.xlsx
```

Note: full `.xlsx` zip files are not expected to be byte-for-byte identical to Apache POI output because zip metadata and document timestamps can vary. Byte-level parity is asserted at the XML writer fixture layer; Phase 0 interoperability is asserted by Apache POI successfully reading dotnet-poi output.

### Phase 1 Verification

Phase 1 is complete for the minimal readable surface: open a simple `.xlsx`, restore sheets/rows/cells, and read string and numeric cell values.

Verification currently covers:

- unit tests for C# write → C# read round-trips
- Java interop in the read direction: Apache POI writes an `.xlsx`, then dotnet-poi reads and asserts the cell values
- Java interop in the write direction remains green: dotnet-poi writes an `.xlsx`, then Apache POI reads it
- a runnable interoperability example under `examples/Phase1InteropExample`

Commands:

```bash
mvn test -f tests/DotnetPoi.Interop.Tests/java/pom.xml -Dtest=WriteForDotnetTest
dotnet test tests/DotnetPoi.XSSF.Tests/DotnetPoi.XSSF.Tests.csproj --no-restore
dotnet test tests/DotnetPoi.Interop.Tests/cs/DotnetPoi.Interop.Tests.csproj --no-restore
mvn test -f tests/DotnetPoi.Interop.Tests/java/pom.xml -Dtest=ReadFromDotnetTest
dotnet run --project examples/Phase1InteropExample/Phase1InteropExample.csproj
```

The example checks:

```text
examples/output/phase1-dotnet-poi-write.xlsx
tests/DotnetPoi.Interop.Tests/fixtures/from-poi/phase1-basic.xlsx
```

Scope note: Phase 1 covers simple `.xlsx` files with shared-string cells, numeric cells, explicit zero values, sparse cells, and multiple sheets. Formulas, styles, dates, booleans, images, and rich text remain later-phase work unless otherwise noted by tests.

### Phase 2 Verification

Phase 2 is complete for the initial practical style subset: fonts, indexed colors, foreground fills, basic borders, built-in number formats, custom number formats, and cell style references.

Verification currently covers:

- unit tests for style table generation and style readback
- Java interop in the write direction: Apache POI reads the dotnet-poi styled workbook fixture
- the XML parity writer tests remain green

Scope note: Phase 2 does not claim full style byte-level parity for every Apache POI style feature. The supported subset is semantically interoperable with Apache POI; additional style properties are tracked as Phase 2.x work.

### Phase 2.5 Verification

Phase 2.5 has its first completed slice: embedding PNG/JPEG-style image parts in `.xlsx` output through POI-compatible APIs.

Implemented surface:

- `XSSFWorkbook.addPicture(...)`
- `XSSFWorkbook.getAllPictures()`
- `XSSFCreationHelper.createClientAnchor()`
- `XSSFSheet.createDrawingPatriarch()`
- `XSSFDrawing.createPicture(...)`
- `XSSFClientAnchor`, `XSSFDrawing`, `XSSFPicture`, and `XSSFPictureData`

Verification currently covers:

- unit tests for generated media, drawing, worksheet relationship, drawing relationship, and content type parts
- C# write → C# read picture data round-trip
- Java interop in the write direction: dotnet-poi writes an image workbook, then Apache POI reads the image data and drawing shape
- a runnable example under `examples/Phase25ImagesExample`

Commands:

```bash
dotnet test tests/DotnetPoi.XSSF.Tests/DotnetPoi.XSSF.Tests.csproj --no-restore
dotnet test tests/DotnetPoi.Interop.Tests/cs/DotnetPoi.Interop.Tests.csproj --no-restore --filter Category=WriteForPoi
mvn test -f tests/DotnetPoi.Interop.Tests/java/pom.xml -Dtest=ReadFromDotnetTest
dotnet run --project examples/Phase25ImagesExample/Phase25ImagesExample.csproj
```

The example writes:

```text
examples/output/phase2_5-images-example.xlsx
```

Byte-level note: the low-level XML writer parity fixture suite is still the source of truth for POI/XMLBeans byte-level XML behavior and remains green. Full workbook ZIP byte parity is not claimed because ZIP metadata and document timestamps vary. Full byte-for-byte parity for every new drawing-related package part is also not claimed yet; current Phase 2.5 verification is semantic package interoperability with Apache POI plus continued `PoiXmlWriter` fixture parity.

### Phase 3.2 Verification

Phase 3.2 is complete for the initial docx surface: create paragraphs with text and inline images, apply bold/italic formatting, set image rotation, save as `.docx`, and read the result back.

Implemented classes: `XWPFDocument`, `XWPFParagraph`, `XWPFRun`, `XWPFPicture`, `XWPFPictureData`.

Verification currently covers:

- unit tests for write/read round-trips (text, bold, italic, inline image, rotation, deduplication)
- Java interop in the write direction: dotnet-poi writes a `.docx`, then Apache POI (XWPF) reads and asserts paragraph text, formatting, picture data, and rotation
- a runnable example under `examples/Phase32DocxExample`

Commands:

```bash
dotnet test tests/DotnetPoi.XWPF.Tests/ --no-restore
dotnet test tests/DotnetPoi.Interop.Tests/cs/ --no-restore --filter Category=WriteForPoi
mvn test -f tests/DotnetPoi.Interop.Tests/java/pom.xml -Dtest=ReadFromDotnetTest
dotnet run --project examples/Phase32DocxExample/Phase32DocxExample.csproj
```

The example writes:

```text
examples/output/phase3_2-docx-example.docx
```

Scope note: Phase 3.2 covers the core XWPF write/read surface — paragraphs, runs, bold/italic, and inline images with rotation. Table support, headers/footers, styles, and numbering are later-phase work.

### Phase 3.3 Verification

Phase 3.3 is complete for the initial pptx surface: create slides with embedded images, set position and size in EMU, apply rotation and horizontal/vertical flip, save as `.pptx`, and read the result back.

Implemented classes: `XMLSlideShow`, `XSLFSlide`, `XSLFPictureShape`, `XSLFPictureData`.

Verification currently covers:

- 18 unit tests for write/read round-trips (slide count, picture bytes, anchor, rotation, flip, deduplication)
- Java interop in the write direction: dotnet-poi writes a `.pptx`, then Apache POI (XSLF) reads and asserts slide count, picture bytes, and rotation (90°)
- a runnable example under `examples/Phase33PptxExample`

Commands:

```bash
dotnet test tests/DotnetPoi.XSLF.Tests/ --no-restore
dotnet test tests/DotnetPoi.Interop.Tests/cs/ --no-restore --filter Category=WriteForPoi
mvn test -f tests/DotnetPoi.Interop.Tests/java/pom.xml -Dtest=ReadFromDotnetTest
dotnet run --project examples/Phase33PptxExample/Phase33PptxExample.csproj
```

The example writes:

```text
examples/output/phase3_3-pptx-example.pptx
```

**Byte-level XML parity note**: All XML output goes through `PoiXmlWriter` — the same foundation used for XSSF and XWPF. This ensures `<tag/>` empty-element style (no space before slash), correct XML declaration format, and no BOM. Full PPTX-specific XML fixture parity tests (analogous to the `xml-parity/` fixtures for XSSF) are not yet implemented; the current guarantee is semantic interoperability with Apache POI Java confirmed by the `ReadFromDotnetTest` suite.

Scope note: Phase 3.3 covers the core XSLF write/read surface — slides, picture shapes with anchor/rotation/flip, and picture data deduplication. Text boxes, animations, charts, and themes are later-phase work.

### Phase 3.4 Verification

Phase 3.4 is complete for the initial OOXML Agile encryption surface: write a password-protected `.xlsx`, wrap it as an OLE2 compound file with `EncryptionInfo` and `EncryptedPackage`, decrypt it through dotnet-poi, and validate the encrypted fixture with Apache POI Java and Microsoft Excel.

Implemented surface:

- `XSSFWorkbook.writeEncrypted(Stream, password)`
- `EncryptionInfo(EncryptionMode.agile)`
- `EncryptionInfo(Stream)` for reading the OLE2 wrapper
- `Encryptor.confirmPassword(...)` and `Encryptor.encryptPackage(...)`
- `Decryptor.verifyPassword(...)` and `Decryptor.getData()`
- in-repo CFB writer/reader for the Agile two-stream POIFS wrapper, including mini FAT support

Verification currently covers:

- 8 unit tests for Agile payload encryption/decryption across chunk and padding boundaries
- C# write fixture: dotnet-poi writes an Agile-encrypted `.xlsx`
- Java interop in the write direction: Apache POI reads and decrypts the dotnet-poi encrypted fixture
- Microsoft Excel manual verification: the generated encrypted workbook opens with password `f`
- a runnable example under `examples/Phase34AgileEncryptionExample`

Commands:

```bash
dotnet test tests/DotnetPoi.POIFS.Tests/DotnetPoi.POIFS.Tests.csproj
dotnet test tests/DotnetPoi.Interop.Tests/cs/DotnetPoi.Interop.Tests.csproj --filter "FullyQualifiedName~Write_AgileEncryptedWorkbook_CreatesFixtureForPoi"
mvn test -f tests/DotnetPoi.Interop.Tests/java/pom.xml -Dtest=ReadFromDotnetTest#readPhase34AgileEncryptedWorkbook
dotnet run --project examples/Phase34AgileEncryptionExample/Phase34AgileEncryptionExample.csproj
```

The example writes:

```text
examples/output/phase3_4-agile-encrypted-example.xlsx
```

Password:

```text
f
```

**Byte-level XML parity note**: Phase 3.4 keeps the existing POI/XMLBeans byte-level XML fixture guarantee for normal OOXML parts. The Agile `EncryptionInfo` XML is intentionally generated in the known Excel-compatible shape documented in `PHASE_3_4_AGILE_ENCRYPTION_NOTES.md`; full encrypted OLE2 file byte-for-byte parity with Apache POI is not claimed because salts, keys, encrypted bytes, sector layout, and package metadata vary. The current guarantee is semantic compatibility with Apache POI and Excel for the implemented AES-128/SHA1 Agile path.

Scope note: Phase 3.4 currently covers the default AES-128/SHA1 Agile encryption path for OOXML `.xlsx`. AES-192/AES-256, SHA-256+, broader POIFS/HSSF, and general-purpose OLE2 document authoring remain later work.

### Phase 4 Verification

Phase 4 is in progress. The first HSSF `.xls` bootstrap slice is present: create/read/write simple BIFF8 workbooks through `HSSFWorkbook`, `HSSFSheet`, `HSSFRow`, and `HSSFCell`.

Implemented HSSF bootstrap surface:

- OLE2 `.xls` container read/write through the in-repo POIFS CFB reader/writer
- BIFF8 Workbook stream read/write for `BoundSheet8`, `SST`/`LabelSST`, `Number`, `RK`, `BoolErr`, `Blank`, `Dimensions`, `Window2`, and `Selection`
- string, numeric, boolean, blank, and error cells
- sheet creation and basic row/cell access

Verification currently covers:

- C# HSSF write → C# HSSF read round-trip
- dotnet-poi reads a POI `.xls` sample from `poi/test-data/spreadsheet`
- Direction A: Apache POI Java writes `.xls` → dotnet-poi reads string/numeric/boolean values
- Direction B: dotnet-poi writes `.xls` → Apache POI Java reads string/numeric/boolean values
- a runnable example under `examples/Phase4HssfXlsExample`

Commands:

```bash
dotnet test tests/DotnetPoi.HSSF.Tests/DotnetPoi.HSSF.Tests.csproj
mvn test -f tests/DotnetPoi.Interop.Tests/java/pom.xml -Dtest=WriteForDotnetTest#writePhase6BasicHssfWorkbook
dotnet test tests/DotnetPoi.Interop.Tests/cs/DotnetPoi.Interop.Tests.csproj --filter "FullyQualifiedName~Read_HssfWorkbook"
dotnet test tests/DotnetPoi.Interop.Tests/cs/DotnetPoi.Interop.Tests.csproj --filter "FullyQualifiedName~Write_HssfWorkbook"
mvn test -f tests/DotnetPoi.Interop.Tests/java/pom.xml -Dtest=ReadFromDotnetTest#readPhase6BasicHssfWorkbook
dotnet run --project examples/Phase4HssfXlsExample/Phase4HssfXlsExample.csproj
```

The example writes:

```text
examples/output/phase4-hssf-xls-example.xls
```

Scope note: Phase 4 is not complete. Formula tokenization/evaluation, styles, pictures, charts, full POIFS directory preservation, broader BIFF records, and full `.xls` parity remain backlog work.

### Phase 7 Verification

Phase 7 (gleaning) adds full cell type coverage to the xlsx read/write path — filling the gap that caused crashes when reading Excel files with formula cached values, boolean cells, or error cells.

Ported from:
- `XSSFCell.getBaseCellType()` — `STCellType` → `CellType` mapping
- `XSSFCell.getCachedFormulaResultType()` — cached formula result type
- `XSSFCell.getBooleanCellValue()` — `"1"` = true, `"0"` = false
- `XSSFCell.getErrorCellString()` — raw OOXML error string
- `CryptoFunctions.generateKey()` — attribute write order fix (`r t s`, matching XMLBeans output)

Cell types now supported:

| OOXML `t` | Read as | Write as |
|---|---|---|
| absent / `"n"` with `<v>` | `Numeric` | no `t` attr ✅ |
| absent / `"n"` without `<v>` | `Blank` | skipped ✅ |
| `"s"` | `String` (shared) | `t="s"` ✅ |
| `"str"` | `Formula → String` | formula written as cached String ✅ |
| `"b"` | `Boolean` | `t="b"`, `<v>1</v>` or `<v>0</v>` ✅ |
| `"e"` | `Error` | `t="e"`, `<v>#DIV/0!</v>` ✅ |
| + `<f>` | `Formula` (read formula and cached `<v>`) | `<f>` plus optional cached `<v>` ✅ |

Verification currently covers:

- Direction A: Apache POI Java writes formula/boolean/error xlsx → dotnet-poi reads all types correctly
- Direction B: dotnet-poi writes boolean cells → Apache POI Java reads and verifies type and value
- C# write → C# read round-trip for all cell types
- a runnable example under `examples/Phase7CellTypesExample`

Commands:

```bash
mvn test -f tests/DotnetPoi.Interop.Tests/java/pom.xml -Dtest=WriteForDotnetTest
dotnet test tests/DotnetPoi.Interop.Tests/cs/ --filter Category=ReadFromPoi
dotnet test tests/DotnetPoi.Interop.Tests/cs/ --filter Category=WriteForPoi
mvn test -f tests/DotnetPoi.Interop.Tests/java/pom.xml -Dtest=ReadFromDotnetTest
dotnet run --project examples/Phase7CellTypesExample/Phase7CellTypesExample.csproj
```

**Byte-level XML parity**: Attribute write order is now `r t s` (matching POI/XMLBeans). The `t="n"` default and `s="0"` default are omitted (OOXML allows this), which is a remaining minor difference from POI. The `PoiXmlWriter` gate tests remain green.

### Phase 5 Verification

Phase 5 has a completed representative slice, not the full Apache POI formula engine. The implemented scope covers formula text save/read, asking Excel to recalculate on open, and evaluating common workbook-local formulas.

Implemented surface:

- `ICell.setCellFormula(...)` / `ICell.getCellFormula()`
- formula XML read/write through `<f>` with cached `<v>` values
- `IWorkbook.setForceFormulaRecalculation(...)` / `getForceFormulaRecalculation()`
- `ICreationHelper.createFormulaEvaluator()`
- `IFormulaEvaluator.evaluate(...)`, `evaluateFormulaCell(...)`, `evaluateAll()`, `evaluateInCell(...)`
- `CellValue`

Representative evaluator support:

- arithmetic: `+`, `-`, `*`, `/`
- cell references and ranges: `A1`, `A1:C1`
- functions: `SUM`, `AVERAGE`, `MIN`, `MAX`, `COUNT`, `CONCATENATE`
- string concatenation with `&`

Verification currently covers:

- C# formula write/read round-trip with cached numeric and string results
- C# `calcPr fullCalcOnLoad` round-trip
- C# formula evaluator unit tests for arithmetic, ranges, aggregate functions, and string formulas
- Direction A: Apache POI writes formulas / recalculation flag → dotnet-poi reads formula text and cached values
- Direction B: dotnet-poi writes evaluated formulas → Apache POI reads formula text and cached values
- a runnable example under `examples/Phase5FormulaEvaluatorExample`

Commands:

```bash
dotnet test tests/DotnetPoi.XSSF.Tests/DotnetPoi.XSSF.Tests.csproj --no-restore
dotnet test tests/DotnetPoi.SS.Tests/DotnetPoi.SS.Tests.csproj --no-restore
dotnet test tests/DotnetPoi.Interop.Tests/cs/DotnetPoi.Interop.Tests.csproj --no-restore --filter Category=ReadFromPoi
dotnet test tests/DotnetPoi.Interop.Tests/cs/DotnetPoi.Interop.Tests.csproj --no-restore --filter Category=WriteForPoi
mvn test -f tests/DotnetPoi.Interop.Tests/java/pom.xml -Dtest=ReadFromDotnetTest
dotnet run --project examples/Phase5FormulaEvaluatorExample/Phase5FormulaEvaluatorExample.csproj
```

Scope note: full Apache POI FormulaEvaluator parity is still partial. External workbook references, shared/array formulas, date functions, lookup functions, financial functions, parser edge cases, and full Excel error semantics remain later Phase 5 work.

---

## Quick Start

> ⚠️ NuGet package not yet published. Use a project reference or clone the repository directly.

```bash
git clone --recurse-submodules https://github.com/yourname/dotnet-poi
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
dotnet run --project examples/Phase5FormulaEvaluatorExample/Phase5FormulaEvaluatorExample.csproj
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
├── poi/                        # Apache POI submodule (read-only)
├── examples/                   # Runnable examples (see examples/README.md)
├── src/
│   ├── DotnetPoi.Core/         # ★ NuGet: DotnetPoi.Core (all implementations)
│   │   ├── SS/                 #   Common interfaces, enums, XML writer
│   │   ├── POIFS/              #   OLE2 compound document container
│   │   ├── XSSF/               #   xlsx (Excel 2007+)
│   │   ├── HSSF/               #   xls (Excel 97-2003 / BIFF)
│   │   ├── XWPF/               #   docx (Word 2007+)
│   │   ├── HWPF/               #   doc (Word 97-2003)
│   │   ├── XSLF/               #   pptx (PowerPoint 2007+)
│   │   └── HSLF/               #   ppt (PowerPoint 97-2003)
│   └── DotnetPoi.Formula/      # ★ NuGet: DotnetPoi.Formula (evaluator only)
├── tests/
│   ├── DotnetPoi.Core.Tests/       # Core tests (188) — all formats
│   ├── DotnetPoi.Formula.Tests/    # Formula evaluator tests (10)
│   ├── DotnetPoi.Interop.Tests/   # Java/.NET bidirectional compatibility tests
│   │   ├── java/                #   Maven project (Apache POI dependency)
│   │   ├── fixtures/            #   Files exchanged between Java and C#
│   │   └── *.cs                 #   C# side of interop tests
│   └── test-files/              # Shared test data files (xlsm, docm, pptm, jpg)
├── tools/
│   └── porter/             # Porting progress tracker
├── agents.md               # LLM porting instructions
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

---
---

# dotnet-poi（日本語）

[Apache POI](https://poi.apache.org/) の **非公式** で忠実な .NET 移植です。

[![CI](https://github.com/modeverv/dotnet-poi/actions/workflows/ci.yml/badge.svg)](https://github.com/modeverv/dotnet-poi/actions/workflows/ci.yml)
[![Examples](https://github.com/modeverv/dotnet-poi/actions/workflows/examples.yml/badge.svg)](https://github.com/modeverv/dotnet-poi/actions/workflows/examples.yml)
[![XML Parity](https://github.com/modeverv/dotnet-poi/actions/workflows/xml-parity-fixtures.yml/badge.svg)](https://github.com/modeverv/dotnet-poi/actions/workflows/xml-parity-fixtures.yml)

## 理念

- 🔱 上流の Apache POI に最大限準拠 — 独自実装ではなく追従する
- 🤖 LLM の支援によりクラス単位で移植し、テストコードも同時に作成
- 💸 永久に無料。EULA なし。メンテナンス費なし。例外なし。
- 📖 Apache POI をソースの正典として git submodule で参照
- ⚠️ Apache Software Foundation とは一切関係ありません（非公式）

---

## 対応状況

### 現在のフェーズ: Phase 4 — HSSF xls bootstrap 対応中

| Phase | 内容 | バージョン目標 | 状態 |
|---|---|---|---|
| **-1** | **XML 出力挙動の統一（Java vs .NET）** | **—** | ✅ 完了 |
| **0** | **xlsx 書き出し（文字・数値）** | **v0.1** | ✅ 完了 |
| **1** | **xlsx 読み込み（文字・数値）** | **v0.2** | ✅ 完了 |
| **2** | **スタイル・書式（フォント・色・罫線）** | **v0.3** | ✅ 完了 |
| **2.5** | **画像・図形（XSSFPicture、XSSFDrawing）** | **v0.35** | ✅ 完了 |
| **3** | **SS 共通インターフェース（IWorkbook / ISheet / IRow / ICell）** | **v0.4** | ✅ 完了 |
| **3.1** | **画像の回転（XSSFPicture.setRotation/getRotation）** | **—** | ✅ 完了 |
| **3.2** | **docx 対応（XWPF — テキスト・画像・回転）** | **—** | ✅ 完了 |
| **3.3** | **pptx 対応（XSLF — スライド・画像・回転・フリップ）** | **—** | ✅ 完了 |
| **3.4** | **AGILE 暗号化（XSSF write / read decrypt path）** | **—** | ✅ 完了 |
| **7** | **落穂拾い — セルタイプ全対応（Boolean・Formula cached・Error 読み書き）** | **—** | ✅ 完了 |
| **5.1** | **数式の保存・読み込み（`<f>` round-trip）** | **—** | ✅ 完了 |
| **5.2** | **Excel に再計算させる設定（`calcPr fullCalcOnLoad`）** | **—** | ✅ 完了 |
| **5.3** | **代表的な関数の評価** | **—** | ✅ 完了 |
| 4 | POIFS + HSSF（xls 読み書き） | v0.5 | 🚧 HSSF bootstrap |
| 5 | FormulaEvaluator 完全互換 | v1.0 | 🚧 一部対応 |
| **6** | **HWPF (.doc テキスト読み込み) + HSLF (.ppt スライド/テキスト読み込み)** | **v1.x** | ✅ 完了（読み込み専用 MVP）|

注意: ライブラリ内での数式の評価と数式結果値の設定は当面オミットします。数式はテキストとして保持し、キャッシュ値がある場合のみ取り扱います。

POIFS を「フル実装」と見なすための最低到達ライン（HWPF/HSLF の土台用）:

- OLE2 ヘッダ、FAT、mini FAT、DIFAT の読み書き（複数ストリーム対応）
- Directory ツリー（storage/stream エントリ、兄弟順序、タイムスタンプ、CLSID）
- mini stream と cutoff の扱い
- 非連続/大容量ストリームのセクタチェーン検証
- ランダムアクセス read/write（seek, length, overwrite）
- CodePage と UTF-16LE 名称保存ルール
- 未知のストリーム/プロパティを壊さずに保持
- テストフィクスチャ:
  - mini/regular 混在ストリームの round-trip
  - POI の comparator に一致するディレクトリ順序
  - Java POI 相互運用（POI が dotnet-poi CFB を読める／その逆）

### Phase 0 クラス別進捗

| クラス | 移植 | テスト | 備考 |
|---|---|---|---|
| `XSSFWorkbook` | ✅ | ✅ | 最小 `.xlsx` パッケージ書き出し |
| `XSSFSheet` | ✅ | ✅ | 最小のシート作成・行アクセス |
| `XSSFRow` | ✅ | ✅ | 最小の行作成・セルアクセス |
| `XSSFCell` | ✅ | ✅ | 文字列・数値・真偽値・エラー・cached value 付き数式 |
| `XSSFCreationHelper` | ✅ | ✅ | data format、anchor、formula evaluator |

凡例: ✅ 完了 / 🚧 進行中 / ⬜ 未着手

### Byte-level parity の対象範囲

現在の XML byte-level parity は、`tests/DotnetPoi.Interop.Tests/fixtures/xml-parity/` 以下の Apache POI/XMLBeans 生成 XSSF `.xlsx` XML part に対する `PoiXmlWriter` fixture 層で保証しています。

これは `.doc`、`.docx`、`.xls`、`.xlsx`、`.ppt`、`.pptx`、`.docm`、`.xlsm`、`.pptm` のファイル全体が POI と byte-for-byte 一致するという意味ではありません。

- `.xls`、`.doc`、`.ppt` は OLE2/binary format であり、XML package format ではありません。
- `.xlsx` は XML writer 挙動を fixture で固定していますが、ZIP metadata や timestamp が変わるため、ファイル全体の byte parity は主張していません。
- `.docx`、`.pptx`、`.docm`、`.xlsm`、`.pptm` は実装済み範囲で semantic round-trip と Apache POI interop を確認していますが、XSSF の `xml-parity/` と同等の format 固有 XML fixture byte-parity suite はまだありません。
- Agile 暗号化は、暗号化前の通常 OOXML part について同じ XML writer 基盤を使います。ただし暗号化済み OLE2 output は salt、key、encrypted payload、sector layout、package metadata が変わるため Apache POI との byte-for-byte 一致は主張しません。Agile `EncryptionInfo` XML は Excel/POI 互換の既知 shape で生成しており、POI と byte-identical として assert しているものではありません。

### Phase -1 基盤

Phase -1 は完了しました。Apache POI/XMLBeans の OOXML 出力にバイト列レベルで寄せるための基盤として `PoiXmlWriter` を追加しています。

固定済みの内容:

- `tests/DotnetPoi.Interop.Tests/fixtures/xml-parity/` 以下の Java Apache POI 生成 fixture
- XML 宣言、空要素、属性順、namespace 順、ゼロ値・デフォルト値属性、要素順、空白、数値表現の byte-level fixture 比較
- production code が `PoiXmlWriter` を迂回して `XmlWriter`、`XDocument`、`XElement`、`XmlDocument`、`XmlSerializer` などを直接使った場合に落ちるゲートテスト

今後の作業では、この XML parity テストが通っていることを確認します。

```bash
dotnet test tests/DotnetPoi.SS.Tests/DotnetPoi.SS.Tests.csproj --filter PoiXmlWriter
```

### Phase 0 検証

Phase 0 は、ワークブック・シート・行・セルを作成し、文字列と数値を書き込み、`.xlsx` として保存できる最初の書き出し面として完了しています。

現在の検証内容:

- Phase 0 XSSF API と生成 OOXML パーツの unit test
- Apache POI/XMLBeans から採取した fixture に対する `PoiXmlWriter` の XML byte-level parity test
- dotnet-poi が `.xlsx` を書き、Apache POI(Java) が読み取ってセル値を検証する相互運用テスト
- `examples/Phase0WriteExample` の実行サンプル

確認コマンド:

```bash
dotnet test tests/DotnetPoi.XSSF.Tests/DotnetPoi.XSSF.Tests.csproj --no-restore
dotnet test tests/DotnetPoi.SS.Tests/DotnetPoi.SS.Tests.csproj --no-restore
dotnet test tests/DotnetPoi.Interop.Tests/cs/DotnetPoi.Interop.Tests.csproj --no-restore --filter Category=WriteForPoi
mvn test -f tests/DotnetPoi.Interop.Tests/java/pom.xml -Dtest=ReadFromDotnetTest
dotnet run --project examples/Phase0WriteExample/Phase0WriteExample.csproj
```

サンプルの出力先:

```text
examples/output/phase0-write-example.xlsx
```

注意: 完全な `.xlsx` zip ファイルは、zip metadata や document timestamp により Apache POI 出力と完全なバイト列一致にはなりません。バイト列一致は XML writer fixture 層で確認し、Phase 0 の相互運用性は Apache POI が dotnet-poi 出力を読めることで確認しています。

### Phase 1 検証

Phase 1 は、シンプルな `.xlsx` を開き、シート・行・セルを復元し、文字列セルと数値セルを読み取れる最小の読み込み面として完了しています。

現在の検証内容:

- C# 書き出し → C# 読み込みの round-trip unit test
- Apache POI(Java) が `.xlsx` を書き、dotnet-poi が読み取ってセル値を検証する相互運用テスト
- dotnet-poi が `.xlsx` を書き、Apache POI(Java) が読み取る既存の相互運用テストも green
- `examples/Phase1InteropExample` の実行サンプル

確認コマンド:

```bash
mvn test -f tests/DotnetPoi.Interop.Tests/java/pom.xml -Dtest=WriteForDotnetTest
dotnet test tests/DotnetPoi.XSSF.Tests/DotnetPoi.XSSF.Tests.csproj --no-restore
dotnet test tests/DotnetPoi.Interop.Tests/cs/DotnetPoi.Interop.Tests.csproj --no-restore
mvn test -f tests/DotnetPoi.Interop.Tests/java/pom.xml -Dtest=ReadFromDotnetTest
dotnet run --project examples/Phase1InteropExample/Phase1InteropExample.csproj
```

サンプルが確認するファイル:

```text
examples/output/phase1-dotnet-poi-write.xlsx
tests/DotnetPoi.Interop.Tests/fixtures/from-poi/phase1-basic.xlsx
```

範囲の注意: Phase 1 が対象にしているのは shared string の文字列セル、数値セル、明示的なゼロ値、疎なセル、複数シートを含むシンプルな `.xlsx` です。数式、スタイル、日付、真偽値、画像、rich text は、テストで明示されるまでは後続フェーズの対象です。

### Phase 2 検証

Phase 2 は、実用的な最初の style subset として完了しています。対象は font、indexed color、foreground fill、basic border、built-in/custom number format、cell style reference です。

現在の検証内容:

- style table 生成と style readback の unit test
- dotnet-poi が styled workbook を書き、Apache POI(Java) が読み取る相互運用テスト
- XML parity writer test が green のまま維持されていること

注意: Apache POI の全 style feature について byte-level parity を主張する段階ではありません。現時点の保証は、実装済み subset の意味的な相互運用性です。

### Phase 2.5 検証

Phase 2.5 は、最初の完了 slice として `.xlsx` への画像埋め込みに対応しています。

実装済み API:

- `XSSFWorkbook.addPicture(...)`
- `XSSFWorkbook.getAllPictures()`
- `XSSFCreationHelper.createClientAnchor()`
- `XSSFSheet.createDrawingPatriarch()`
- `XSSFDrawing.createPicture(...)`
- `XSSFClientAnchor`、`XSSFDrawing`、`XSSFPicture`、`XSSFPictureData`

現在の検証内容:

- media / drawing / worksheet rels / drawing rels / content types の生成 unit test
- C# 書き出し → C# 読み込みで picture data を round-trip
- dotnet-poi が image workbook を書き、Apache POI(Java) が画像データと drawing shape を読み取る相互運用テスト
- `examples/Phase25ImagesExample` の実行サンプル

確認コマンド:

```bash
dotnet test tests/DotnetPoi.XSSF.Tests/DotnetPoi.XSSF.Tests.csproj --no-restore
dotnet test tests/DotnetPoi.Interop.Tests/cs/DotnetPoi.Interop.Tests.csproj --no-restore --filter Category=WriteForPoi
mvn test -f tests/DotnetPoi.Interop.Tests/java/pom.xml -Dtest=ReadFromDotnetTest
dotnet run --project examples/Phase25ImagesExample/Phase25ImagesExample.csproj
```

サンプルの出力先:

```text
examples/output/phase2_5-images-example.xlsx
```

byte-level に関する注意: Apache POI/XMLBeans の XML byte-level 挙動は、引き続き低レイヤの `PoiXmlWriter` fixture test で固定しており green です。一方で `.xlsx` 全体の zip byte-level 一致は、zip metadata や document timestamp が変わるため主張していません。また drawing 関連の全 package part について完全な byte-for-byte parity を主張する段階でもありません。Phase 2.5 の現時点の保証は、Apache POI との意味的な package 相互運用性と、`PoiXmlWriter` fixture parity の維持です。

### Phase 3.2 検証

Phase 3.2 は最初の docx 書き出し面として完了しています。対象は段落テキスト・インライン画像・bold/italic・画像の回転・docx 保存・読み込みです。

実装済みクラス: `XWPFDocument`、`XWPFParagraph`、`XWPFRun`、`XWPFPicture`、`XWPFPictureData`

現在の検証内容:

- write/read round-trip の unit test（テキスト・書式・インライン画像・回転・重複排除）
- dotnet-poi が `.docx` を書き、Apache POI(Java / XWPF) が段落テキスト・書式・画像データ・回転を読み取る相互運用テスト
- `examples/Phase32DocxExample` の実行サンプル

確認コマンド:

```bash
dotnet test tests/DotnetPoi.XWPF.Tests/ --no-restore
dotnet test tests/DotnetPoi.Interop.Tests/cs/ --no-restore --filter Category=WriteForPoi
mvn test -f tests/DotnetPoi.Interop.Tests/java/pom.xml -Dtest=ReadFromDotnetTest
dotnet run --project examples/Phase32DocxExample/Phase32DocxExample.csproj
```

サンプルの出力先:

```text
examples/output/phase3_2-docx-example.docx
```

範囲の注意: Phase 3.2 が対象にしているのは XWPF の書き出し・読み込みコアサーフェス（段落・ラン・bold/italic・回転付きインライン画像）です。表・ヘッダー/フッター・スタイル・リストは後続フェーズの対象です。

### Phase 3.3 検証

Phase 3.3 は最初の pptx 書き出し面として完了しています。対象はスライド作成・画像埋め込み・アンカー設定（EMU）・回転・水平/垂直フリップ・pptx 保存・読み込みです。

実装済みクラス: `XMLSlideShow`、`XSLFSlide`、`XSLFPictureShape`、`XSLFPictureData`

現在の検証内容:

- write/read round-trip の 18 unit test（スライド数・画像バイト列・アンカー・回転・フリップ・重複排除）
- dotnet-poi が `.pptx` を書き、Apache POI(Java / XSLF) がスライド数・画像バイト列・回転（90°）を読み取る相互運用テスト
- `examples/Phase33PptxExample` の実行サンプル

確認コマンド:

```bash
dotnet test tests/DotnetPoi.XSLF.Tests/ --no-restore
dotnet test tests/DotnetPoi.Interop.Tests/cs/ --no-restore --filter Category=WriteForPoi
mvn test -f tests/DotnetPoi.Interop.Tests/java/pom.xml -Dtest=ReadFromDotnetTest
dotnet run --project examples/Phase33PptxExample/Phase33PptxExample.csproj
```

サンプルの出力先:

```text
examples/output/phase3_3-pptx-example.pptx
```

**byte-level XML parity に関する注意**: 全 XML 出力は `PoiXmlWriter` を経由しており、XSSF・XWPF と同等の基盤です（空要素スタイル `<tag/>`、XML 宣言形式、BOM なし）。PPTX 固有の XML fixture parity テスト（XSSF の `xml-parity/` fixtures に相当するもの）は未整備です。現時点の保証は `ReadFromDotnetTest` で確認した Apache POI Java との意味的な相互運用性です。

範囲の注意: Phase 3.3 が対象にしているのは XSLF の書き出し・読み込みコアサーフェス（スライド・アンカー付き画像・回転・フリップ・重複排除）です。テキストボックス・アニメーション・グラフ・テーマは後続フェーズの対象です。

### Phase 3.4 検証

Phase 3.4 は最初の OOXML Agile encryption surface として完了しています。対象は password protected `.xlsx` の書き出し、`EncryptionInfo` / `EncryptedPackage` を持つ OLE2 compound file へのラップ、dotnet-poi での復号、Apache POI Java と Microsoft Excel での読み込み確認です。

実装済み API:

- `XSSFWorkbook.writeEncrypted(Stream, password)`
- `EncryptionInfo(EncryptionMode.agile)`
- `EncryptionInfo(Stream)` による OLE2 wrapper 読み込み
- `Encryptor.confirmPassword(...)` / `Encryptor.encryptPackage(...)`
- `Decryptor.verifyPassword(...)` / `Decryptor.getData()`
- Agile の 2 stream POIFS wrapper 用 in-repo CFB writer/reader（mini FAT 対応）

現在の検証内容:

- chunk / padding 境界を含む Agile payload 暗号化・復号 unit test 8 件
- dotnet-poi が Agile encrypted `.xlsx` fixture を書く C# interop test
- Apache POI(Java) が dotnet-poi 生成の暗号化 fixture を復号・読み込みする相互運用テスト
- Microsoft Excel で password `f` による手動オープン確認
- `examples/Phase34AgileEncryptionExample` の実行サンプル

確認コマンド:

```bash
dotnet test tests/DotnetPoi.POIFS.Tests/DotnetPoi.POIFS.Tests.csproj
dotnet test tests/DotnetPoi.Interop.Tests/cs/DotnetPoi.Interop.Tests.csproj --filter "FullyQualifiedName~Write_AgileEncryptedWorkbook_CreatesFixtureForPoi"
mvn test -f tests/DotnetPoi.Interop.Tests/java/pom.xml -Dtest=ReadFromDotnetTest#readPhase34AgileEncryptedWorkbook
dotnet run --project examples/Phase34AgileEncryptionExample/Phase34AgileEncryptionExample.csproj
```

サンプルの出力先:

```text
examples/output/phase3_4-agile-encrypted-example.xlsx
```

password:

```text
f
```

**byte-level XML parity に関する注意**: 通常の OOXML part については、既存の Apache POI/XMLBeans byte-level XML fixture guarantee を維持しています。Agile の `EncryptionInfo` XML は `PHASE_3_4_AGILE_ENCRYPTION_NOTES.md` に記録した Excel-compatible shape で意図的に生成しています。暗号化済み OLE2 file 全体の Apache POI との byte-for-byte parity は、salt / key / encrypted bytes / sector layout / package metadata が変わるため主張していません。現時点の保証は、実装済み AES-128/SHA1 Agile path の Apache POI と Excel との意味的な互換性です。

範囲の注意: Phase 3.4 が対象にしているのは OOXML `.xlsx` の default AES-128/SHA1 Agile encryption path です。AES-192/AES-256、SHA-256+、より広い POIFS/HSSF、汎用 OLE2 document authoring は後続フェーズの対象です。

### Phase 4 検証

Phase 4 は対応中です。最初の HSSF `.xls` bootstrap slice として、`HSSFWorkbook`、`HSSFSheet`、`HSSFRow`、`HSSFCell` によるシンプルな BIFF8 workbook の作成・読み込み・書き出しに対応しています。

実装済み HSSF bootstrap surface:

- in-repo POIFS CFB reader/writer による OLE2 `.xls` container read/write
- `BoundSheet8`、`SST`/`LabelSST`、`Number`、`RK`、`BoolErr`、`Blank`、`Dimensions`、`Window2`、`Selection` の BIFF8 Workbook stream read/write
- string / numeric / boolean / blank / error cell
- sheet creation と基本的な row/cell access

現在の検証内容:

- C# HSSF write → C# HSSF read round-trip
- `poi/test-data/spreadsheet` の POI `.xls` sample 読み込み
- A 方向: Apache POI Java が `.xls` を書き出し → dotnet-poi が string/numeric/boolean value を読み込み
- B 方向: dotnet-poi が `.xls` を書き出し → Apache POI Java が string/numeric/boolean value を読み込み
- `examples/Phase4HssfXlsExample` の実行サンプル

確認コマンド:

```bash
dotnet test tests/DotnetPoi.HSSF.Tests/DotnetPoi.HSSF.Tests.csproj
mvn test -f tests/DotnetPoi.Interop.Tests/java/pom.xml -Dtest=WriteForDotnetTest#writePhase6BasicHssfWorkbook
dotnet test tests/DotnetPoi.Interop.Tests/cs/DotnetPoi.Interop.Tests.csproj --filter "FullyQualifiedName~Read_HssfWorkbook"
dotnet test tests/DotnetPoi.Interop.Tests/cs/DotnetPoi.Interop.Tests.csproj --filter "FullyQualifiedName~Write_HssfWorkbook"
mvn test -f tests/DotnetPoi.Interop.Tests/java/pom.xml -Dtest=ReadFromDotnetTest#readPhase6BasicHssfWorkbook
dotnet run --project examples/Phase4HssfXlsExample/Phase4HssfXlsExample.csproj
```

サンプルの出力:

```text
examples/output/phase4-hssf-xls-example.xls
```

範囲の注意: Phase 4 は完了ではありません。Formula tokenization/evaluation、style、picture、chart、full POIFS directory preservation、より広い BIFF record、完全な `.xls` parity は backlog です。

### Phase 7 検証

Phase 7 (gleaning) は xlsx 読み書きパスに全セルタイプ対応を追加します。これにより、数式キャッシュ値、真偽値、エラーセルを含む Excel ファイルの読み込み時クラッシュの原因となっていたギャップが埋まります。

移植元:
- `XSSFCell.getBaseCellType()` — `STCellType` → `CellType` マッピング
- `XSSFCell.getCachedFormulaResultType()` — キャッシュ数式結果タイプ
- `XSSFCell.getBooleanCellValue()` — `"1"` = true, `"0"` = false
- `XSSFCell.getErrorCellString()` — 生の OOXML エラー文字列
- `CryptoFunctions.generateKey()` — 属性書き込み順修正 (`r t s`, XMLBeans 出力に合わせる)

現在サポートされているセルタイプ:

| OOXML `t` | 読み込み時 | 書き込み時 |
|---|---|---|
| absent / `"n"` with `<v>` | `Numeric` | no `t` attr ✅ |
| absent / `"n"` without `<v>` | `Blank` | skipped ✅ |
| `"s"` | `String` (shared) | `t="s"` ✅ |
| `"str"` | `Formula → String` | formula written as cached String ✅ |
| `"b"` | `Boolean` | `t="b"`, `<v>1</v>` or `<v>0</v>` ✅ |
| `"e"` | `Error` | `t="e"`, `<v>#DIV/0!</v>` ✅ |
| + `<f>` | `Formula` (数式とキャッシュ `<v>` の読み込み) | `<f>` plus optional cached `<v>` ✅ |

現在の検証内容:

- A 方向: Apache POI Java が数式/真偽値/エラー xlsx を書き出し → dotnet-poi が全タイプを正しく読み込み
- B 方向: dotnet-poi が真偽値セルを書き出し → Apache POI Java がタイプと値を検証
- C# による全セルタイプの読み書き round-trip
- `examples/Phase7CellTypesExample` の実行サンプル

確認コマンド:

```bash
mvn test -f tests/DotnetPoi.Interop.Tests/java/pom.xml -Dtest=WriteForDotnetTest
dotnet test tests/DotnetPoi.Interop.Tests/cs/ --filter Category=ReadFromPoi
dotnet test tests/DotnetPoi.Interop.Tests/cs/ --filter Category=WriteForPoi
mvn test -f tests/DotnetPoi.Interop.Tests/java/pom.xml -Dtest=ReadFromDotnetTest
dotnet run --project examples/Phase7CellTypesExample/Phase7CellTypesExample.csproj
```

**バイト単位の XML parity**: 属性の書き込み順は `r t s` になり（POI/XMLBeans に合わせる）、`t="n"` デフォルトと `s="0"` デフォルトは省略されます（OOXML では許容される）。これは POI との残りのマイナーな違いです。`PoiXmlWriter` ゲートテストは green のままです。

### Phase 5 検証

Phase 5 は代表的な subset として完了している範囲があります。Apache POI の FormulaEvaluator 全体の完全互換ではありませんが、数式文字列の保存・読み込み、Excel に再計算させる設定、代表的な workbook-local formula の評価に対応しています。

実装済み API:

- `ICell.setCellFormula(...)` / `ICell.getCellFormula()`
- `<f>` と cached `<v>` の read/write
- `IWorkbook.setForceFormulaRecalculation(...)` / `getForceFormulaRecalculation()`
- `ICreationHelper.createFormulaEvaluator()`
- `IFormulaEvaluator.evaluate(...)` / `evaluateFormulaCell(...)` / `evaluateAll()` / `evaluateInCell(...)`
- `CellValue`

対応済み evaluator subset:

- 四則演算: `+`, `-`, `*`, `/`
- セル参照・範囲参照: `A1`, `A1:C1`
- 関数: `SUM`, `AVERAGE`, `MIN`, `MAX`, `COUNT`, `CONCATENATE`
- `&` による文字列結合

確認コマンド:

```bash
dotnet test tests/DotnetPoi.XSSF.Tests/DotnetPoi.XSSF.Tests.csproj --no-restore
dotnet test tests/DotnetPoi.SS.Tests/DotnetPoi.SS.Tests.csproj --no-restore
dotnet test tests/DotnetPoi.Interop.Tests/cs/DotnetPoi.Interop.Tests.csproj --no-restore --filter Category=ReadFromPoi
dotnet test tests/DotnetPoi.Interop.Tests/cs/DotnetPoi.Interop.Tests.csproj --no-restore --filter Category=WriteForPoi
mvn test -f tests/DotnetPoi.Interop.Tests/java/pom.xml -Dtest=ReadFromDotnetTest
dotnet run --project examples/Phase5FormulaEvaluatorExample/Phase5FormulaEvaluatorExample.csproj
```

範囲の注意: full Apache POI FormulaEvaluator parity はまだ部分的です。external workbook references、shared/array formulas、date functions、lookup functions、financial functions、parser edge cases、全 Excel error semantics は後続の Phase 5 作業です。


---

## クイックスタート

> ⚠️ まだ NuGet パッケージは公開されていません。現時点では project reference か repository clone で利用してください。

```bash
git clone --recurse-submodules https://github.com/yourname/dotnet-poi
cd dotnet-poi
dotnet build
dotnet test
```

### 使用例

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

実行可能な example:

```bash
dotnet run --project examples/Phase0WriteExample/Phase0WriteExample.csproj
dotnet run --project examples/Phase1InteropExample/Phase1InteropExample.csproj
dotnet run --project examples/Phase25ImagesExample/Phase25ImagesExample.csproj
dotnet run --project examples/Phase32DocxExample/Phase32DocxExample.csproj
dotnet run --project examples/Phase33PptxExample/Phase33PptxExample.csproj
dotnet run --project examples/Phase34AgileEncryptionExample/Phase34AgileEncryptionExample.csproj
dotnet run --project examples/Phase5FormulaEvaluatorExample/Phase5FormulaEvaluatorExample.csproj
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
run.setText("dotnet-poi から Hello");
run.setBold(true);

using var fs = new FileStream("output.docx", FileMode.Create);
doc.write(fs);
```

---

## なぜこのプロジェクトが必要か

.NET の Excel ライブラリには構造的な問題があります。

- **NPOI**: xls / xlsx 両対応だがv2.8.0 以降は商用利用に維持費が必要
- **ClosedXML / EPPlus**: xlsx のみ対応、xls（BIFF形式）は扱えない

dotnet-poi は Apache POI という枯れた実装を正典として透過的に移植することで、**実装品質と永続的な無償提供を両立**することを目指します。

---

## 移植方針

Apache POI のソースを `poi/` に git submodule として保持し、**常に原典を参照しながら**クラス単位で移植します。LLM が Java → C# の変換を担い、人間がアーキテクチャ判断と品質検証を行います。

これは同時に「LLM が大規模な知的作業をどこまで担えるか」という実験でもあります。

詳細な移植ルールは [agents.md](./agents.md) を参照してください。

---

## リポジトリ構造

```
dotnet-poi/
├── poi/                        # Apache POI submodule（参照専用）
├── src/
│   ├── DotnetPoi.Core/         # ★ NuGet: DotnetPoi.Core（全実装）
│   │   ├── SS/                 #   共通インターフェース・enum・XML writer
│   │   ├── POIFS/              #   OLE2 compound document container
│   │   ├── XSSF/               #   xlsx（Excel 2007+）
│   │   ├── HSSF/               #   xls（Excel 97-2003 / BIFF）
│   │   ├── XWPF/               #   docx（Word 2007+）
│   │   ├── HWPF/               #   doc（Word 97-2003）
│   │   ├── XSLF/               #   pptx（PowerPoint 2007+）
│   │   └── HSLF/               #   ppt（PowerPoint 97-2003）
│   └── DotnetPoi.Formula/      # ★ NuGet: DotnetPoi.Formula（評価器のみ）
├── tests/
│   ├── DotnetPoi.Core.Tests/       # Core テスト（188）— 全フォーマット
│   ├── DotnetPoi.Formula.Tests/    # Formula 評価器テスト（10）
│   ├── DotnetPoi.Interop.Tests/   # Java/.NET 双方向互換性テスト
│   │   ├── java/                #   Maven プロジェクト（Apache POI 依存）
│   │   ├── fixtures/            #   Java/C# 間でやり取りするファイル
│   │   └── *.cs                 #   C# 側の相互テスト
│   └── test-files/              # 共有テストデータ（xlsm, docm, pptm, jpg）
├── tools/
│   └── porter/             # 移植進捗管理スクリプト
├── agents.md               # LLM への移植指示
└── README.md
```

> **設計の補足:** 全フォーマット実装は `DotnetPoi.Core/` 以下の単一アセンブリとしてコンパイルされます。
> 従来のフォーマット別プロジェクトディレクトリ（`DotnetPoi.XSSF/` など）は個別の `.csproj` としては
> 存在せず、ソースファイルは Core に統合されています。名前空間は従来通りです
> （例: `DotnetPoi.XSSF.UserModel`）。

---

## コントリビュート

個人のライフワークプロジェクトですが、PR・Issue は歓迎します。移植に参加する場合は必ず [agents.md](./agents.md) を読んでから作業してください。

---

## ライセンス

[Apache License 2.0](./LICENSE) — 上流の Apache POI と同じです。

---

## 免責事項

このプロジェクトは Apache Software Foundation および Apache POI プロジェクトとは一切関係ありません。Apache POI は Apache Software Foundation の登録商標です。
