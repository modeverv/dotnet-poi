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
| todo | `poi/poi-ooxml/src/test/java/org/apache/poi/xssf/usermodel/` | Survey round-trip workbook tests | Baseline XSSF package realism | TBD |
| todo | `poi/poi-ooxml/src/test/java/org/apache/poi/xssf/usermodel/` | Survey styles/shared strings tests | Multi-part workbook interactions | TBD |
| todo | `poi/poi-ooxml/src/test/java/org/apache/poi/xssf/usermodel/` | Survey drawings/comments tests | Relationships and drawing package parts | TBD |
| todo | `poi/poi-ooxml/src/test/java/org/apache/poi/xssf/model/` | Survey model-level package tests | Lower-level XML/relationship behavior | TBD |
| todo | `poi/poi-ooxml/src/test/java/org/apache/poi/openxml4j/` | Survey OPC package tests | Package preservation and relationship correctness | TBD |
| todo | `poi/test-data/` | Survey reusable `.xlsx` / `.xlsm` fixtures | Real-world package shapes | TBD |

## Execution Steps

1. Survey upstream POI tests with `rg` and shortlist candidates.
2. Add exact candidate rows to this file.
3. Pick one scenario and build a Java fixture generator case.
4. Generate package and extracted XML fixtures.
5. Add C# semantic compatibility tests.
6. Only then decide whether any finding belongs in `PoiXmlWriter`.

## Notes

- This complements `XMLBEANS_XML_OUTPUT_TODO.md`.
- This work should not resurrect the previous fixture-specific `XSSFWorkbook.WriteWorkbook` behavior.
