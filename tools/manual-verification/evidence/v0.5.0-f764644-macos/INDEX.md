# DotnetPOI v0.5.0-f764644-macos macOS Office Evidence

- Project version: `0.5.0`
- Git revision: `f764644`
- Captured: `2026-05-07 12:44:15 +0900` - `2026-05-07 12:47:27 +0900`
- Environment: macOS Microsoft Office apps
- Excel: `16.108.3`
- Word: `16.108.3`
- PowerPoint: `16.108.3`
- Source root: `tools/manual-verification/generated-documents`
- Overall: `PASS`
- Result counts: `9` pass, `0` missing fixture, `0` permission required, `0` fail

This evidence pass opens each available file in the corresponding macOS Microsoft Office app, treats open/reopen failures as failed cases, captures a screen PNG, and writes this index for GitHub review.

## Matrix

| kind | app | source | encrypted | open | reopen | status | evidence | notes |
|---|---|---|---:|---:|---:|---:|---|---|
| xlsx | excel | tools/manual-verification/generated-documents/manual-simple.xlsx | no | PASS | PASS | PASS | <img src="images/01-xlsx-manual-simple.png" width="320" alt="xlsx macOS Office evidence"> |  |
| xlsm | excel | tools/manual-verification/generated-documents/manual-simple.xlsm | no | PASS | PASS | PASS | <img src="images/02-xlsm-manual-simple.png" width="320" alt="xlsm macOS Office evidence"> |  |
| encrypted xlsx | excel | tools/manual-verification/generated-documents/manual-encrypted.xlsx | yes | PASS | PASS | PASS | <img src="images/03-encrypted-xlsx-manual-encrypted.png" width="320" alt="encrypted xlsx macOS Office evidence"> |  |
| docx | word | tools/manual-verification/generated-documents/manual-simple.docx | no | PASS | PASS | PASS | <img src="images/04-docx-manual-simple.png" width="320" alt="docx macOS Office evidence"> |  |
| docm | word | tools/manual-verification/generated-documents/manual-simple.docm | no | PASS | PASS | PASS | <img src="images/05-docm-manual-simple.png" width="320" alt="docm macOS Office evidence"> |  |
| encrypted docx | word | tools/manual-verification/generated-documents/manual-encrypted.docx | yes | PASS | PASS | PASS | <img src="images/06-encrypted-docx-manual-encrypted.png" width="320" alt="encrypted docx macOS Office evidence"> |  |
| pptx | powerpoint | tools/manual-verification/generated-documents/manual-simple.pptx | no | PASS | PASS | PASS | <img src="images/07-pptx-manual-simple.png" width="320" alt="pptx macOS Office evidence"> |  |
| pptm | powerpoint | tools/manual-verification/generated-documents/manual-simple.pptm | no | PASS | PASS | PASS | <img src="images/08-pptm-manual-simple.png" width="320" alt="pptm macOS Office evidence"> |  |
| encrypted pptx | powerpoint | tools/manual-verification/generated-documents/manual-encrypted.pptx | yes | PASS | PASS | PASS | <img src="images/09-encrypted-pptx-manual-encrypted.png" width="320" alt="encrypted pptx macOS Office evidence"> |  |

## Notes

- `MISSING` means generated manual documents are not present; run `dotnet run --project tools/manual-verification/DocumentGenerator/DocumentGenerator.csproj`.
- The original files are not modified; work copies are written under `workfiles/`.
- Password for generated encrypted files: `f`.
- Screenshots are captured with macOS `screencapture`; run permission/bootstrap mode first if preflight reports `PERMISSION_REQUIRED`.
