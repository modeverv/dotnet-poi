# CHECKPOINT

## 2026-05-05 02:xx JST - Agreed recovery plan for XML parity work

- New direction agreed with user:
  1. Remove the `31e9006 parity` work from production/test code for now, including the fixture-specific XSSF writer changes and associated parity tests/fixtures that forced those changes.
  2. Re-study Java POI XML output behavior and XMLBeans behavior at the correct layer. Add reference code/fixture generators where useful, and write a dedicated Markdown TODO/design file describing observed XMLBeans/POI output patterns, open questions, and the implementation order.
  3. Re-implement the behavior incrementally in `PoiXmlWriter`, with focused failing tests per low-level XML divergence. Keep higher-level `XSSFWorkbook` output POI-model-driven, and preserve unknown/original package parts byte-for-byte where the model does not yet support them.
- Important boundary: XML lexical quirks such as declaration format, empty element form, escaping, attribute order, namespace placement, and whitespace belong in `PoiXmlWriter` or focused helpers. Specific workbook content such as defined names, Office revision GUIDs, local absolute paths, workbook extLst contents, and fixture-specific relationship ordering must not be generalized into `XSSFWorkbook` unless directly backed by the POI model/source behavior.
- Do not commit via LLM.
- Step 1 status: completed in working tree by restoring `src/` and the parity-related `tests/` paths back to `7a4b778` (the parent of `31e9006`). This removes the fixture-specific XSSF writer changes, the added XSSF/XWPF/XSLF parity tests, the expanded Java parity fixture generator, and the extra generated xml-parity fixtures. `dotnet test` passes after the removal.

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
