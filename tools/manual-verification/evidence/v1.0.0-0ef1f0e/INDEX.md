# DotnetPOI v1.0.0-0ef1f0e Linux LibreOffice Evidence

- Project version: `1.0.0`
- Git revision: `0ef1f0e`
- Captured: `2026-05-08 10:04:09 +0900` - `2026-05-08 10:04:15 +0900`
- Environment: Docker service `libreoffice`, container `dotnet-poi-phase11-libreoffice`
- LibreOffice: `LibreOffice 25.8.1.1 580(Build:1)`
- Source root: `tools/manual-verification/generated-documents`
- Overall: `PASS`
- Result counts: `12` pass, `0` missing fixture, `0` fail

This evidence pass opens each available file through LibreOffice UNO, rejects failures/exceptions as a failed case, reopens the work copy, exports a PNG preview, and writes this index for GitHub review.

## Matrix

| kind | source | encrypted | open | reopen | status | evidence | notes |
|---|---|---:|---:|---:|---:|---|---|
| xlsx | tools/manual-verification/generated-documents/manual-simple.xlsx | no | PASS | PASS | PASS | <img src="images/01-xlsx-manual-simple.png" width="320" alt="xlsx evidence"> |  |
| xlsm | tools/manual-verification/generated-documents/manual-simple.xlsm | no | PASS | PASS | PASS | <img src="images/02-xlsm-manual-simple.png" width="320" alt="xlsm evidence"> |  |
| encrypted xlsx | tools/manual-verification/generated-documents/manual-encrypted.xlsx | yes | PASS | PASS | PASS | <img src="images/03-encrypted-xlsx-manual-encrypted.png" width="320" alt="encrypted xlsx evidence"> |  |
| docx | tools/manual-verification/generated-documents/manual-simple.docx | no | PASS | PASS | PASS | <img src="images/04-docx-manual-simple.png" width="320" alt="docx evidence"> |  |
| docm | tools/manual-verification/generated-documents/manual-simple.docm | no | PASS | PASS | PASS | <img src="images/05-docm-manual-simple.png" width="320" alt="docm evidence"> |  |
| encrypted docx | tools/manual-verification/generated-documents/manual-encrypted.docx | yes | PASS | PASS | PASS | <img src="images/06-encrypted-docx-manual-encrypted.png" width="320" alt="encrypted docx evidence"> |  |
| pptx | tools/manual-verification/generated-documents/manual-simple.pptx | no | PASS | PASS | PASS | <img src="images/07-pptx-manual-simple.png" width="320" alt="pptx evidence"> |  |
| pptm | tools/manual-verification/generated-documents/manual-simple.pptm | no | PASS | PASS | PASS | <img src="images/08-pptm-manual-simple.png" width="320" alt="pptm evidence"> |  |
| encrypted pptx | tools/manual-verification/generated-documents/manual-encrypted.pptx | yes | PASS | PASS | PASS | <img src="images/09-encrypted-pptx-manual-encrypted.png" width="320" alt="encrypted pptx evidence"> |  |
| xls | tools/manual-verification/generated-documents/manual-simple.xls | no | PASS | PASS | PASS | <img src="images/10-xls-manual-simple.png" width="320" alt="xls evidence"> |  |
| doc | tools/manual-verification/generated-documents/manual-simple.doc | no | PASS | PASS | PASS | <img src="images/11-doc-manual-simple.png" width="320" alt="doc evidence"> |  |
| ppt | tools/manual-verification/generated-documents/manual-simple.ppt | no | PASS | PASS | PASS | <img src="images/12-ppt-manual-simple.png" width="320" alt="ppt evidence"> |  |

## Notes

- `MISSING` means generated manual documents are not present; run `dotnet run --project tools/manual-verification/DocumentGenerator/DocumentGenerator.csproj`.
- The original files are not modified; work copies are written under `workfiles/`.
- PNG previews are exported by LibreOffice itself, not by the browser screenshot path.
- Password for generated encrypted files: `f`.
