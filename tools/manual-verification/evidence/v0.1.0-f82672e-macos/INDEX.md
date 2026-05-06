# DotnetPOI v0.1.0-f82672e-macos macOS Office Evidence

- Project version: `0.1.0`
- Git revision: `f82672e`
- Captured: `2026-05-06 20:36:02 +0900` - `2026-05-06 20:36:38 +0900`
- Environment: macOS Microsoft Office apps
- Excel: `16.108.3`
- Word: `16.108.3`
- PowerPoint: `16.108.3`
- Source roots: `examples/output`, `tests/test-files`, `tests/DotnetPoi.Interop.Tests/fixtures/from-dotnet-poi`
- Overall: `PASS`
- Result counts: `1` pass, `0` missing fixture, `0` permission required, `0` fail

This evidence pass opens each available file in the corresponding macOS Microsoft Office app, treats open/reopen failures as failed cases, captures a screen PNG, and writes this index for GitHub review.

## Matrix

| kind | app | source | encrypted | open | reopen | status | evidence | notes |
|---|---|---|---:|---:|---:|---:|---|---|
| encrypted xlsm | excel | examples/output/3.xlsm | yes | PASS | PASS | PASS | <img src="images/01-encrypted-xlsm-3.png" width="320" alt="encrypted xlsm macOS Office evidence"> |  |

## Notes

- `MISSING` means this repository does not currently contain a fixture for that requested category.
- The original files are not modified; work copies are written under `workfiles/`.
- Passwords currently known to this harness: `f` for `phase3_4-agile-encrypted-example.xlsx`, `edge-pass` for `edge-encrypted-sparse.xlsx`.
- Screenshots are captured with macOS `screencapture`; run permission/bootstrap mode first if preflight reports `PERMISSION_REQUIRED`.
