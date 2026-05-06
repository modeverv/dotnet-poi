# Phase 11 Manual Verification

This directory contains the release-before manual verification harness for Office and LibreOffice behavior.

Phase 11 is intentionally separate from `tests/`. These checks depend on GUI applications, host OS behavior, Docker images, Office installations, and screenshots, so they are not part of the normal `dotnet test` gate.

## Generate Manual Test Documents

Create a fresh set of simple Office files for manual verification:

```bash
dotnet run --project tools/manual-verification/DocumentGenerator/DocumentGenerator.csproj
```

By default the generator writes to:

```text
tools/manual-verification/generated-documents/
```

It creates:

- `manual-simple.xlsx`
- `manual-simple.xlsm`
- `manual-encrypted.xlsx`
- `manual-simple.pptx`
- `manual-simple.pptm`
- `manual-encrypted.pptx`
- `manual-simple.docx`
- `manual-simple.docm`
- `manual-encrypted.docx`
- `manual-simple.xls`

Encrypted files use password `f`. Generated files are ignored by git; rerun the generator whenever manual verification needs a clean set.

You can pass a custom output directory:

```bash
dotnet run --project tools/manual-verification/DocumentGenerator/DocumentGenerator.csproj -- /tmp/dotnet-poi-manual-files
```

## Linux / LibreOffice

Start the Linux LibreOffice desktop container:

```bash
tools/manual-verification/scripts/run-linux-libreoffice.sh
```

Then open:

- Web desktop: http://localhost:3110
- HTTPS web desktop: https://localhost:3111

The repository is mounted at `/workspace` inside the container. Evidence should be written under:

```text
tools/manual-verification/evidence/
```

Generated evidence is ignored by git.

Run the current Linux automated assist against the generated manual documents, excluding `.xls`:

```bash
tools/manual-verification/scripts/run-linux-manual-check.sh
```

The script opens each supported generated Office file in LibreOffice, stores the work copy, closes it, and reopens it. Encrypted generated files are opened with password `f`. It writes:

- `tools/manual-verification/evidence/linux/session.log`
- `tools/manual-verification/evidence/linux/summary.md`
- `tools/manual-verification/evidence/linux/workfiles/`

Generate versioned Linux evidence in one command:

```bash
tools/manual-verification/scripts/run-linux-evidence.sh
```

This command starts the Docker LibreOffice service, reads the current DotnetPOI version and git revision, opens the generated xlsx/xlsm/encrypted-xlsx/docx/docm/encrypted-docx/pptx/pptm/encrypted-pptx files through LibreOffice UNO, exports PNG previews, and writes a GitHub-readable `INDEX.md` under:

```text
tools/manual-verification/evidence/v<version>-<git-sha>/
```

Generate versioned macOS Microsoft Office evidence in one command:

```bash
tools/manual-verification/scripts/run-macos-office-evidence.sh
```

Before running evidence mode for the first time, run permission/bootstrap mode:

```bash
tools/manual-verification/scripts/run-macos-office-permissions.sh
```

This intentionally triggers or checks macOS permissions for Office automation, System Events, and screen capture. It writes its result to:

```text
tools/manual-verification/evidence/macos-permissions/STATUS.md
```

Evidence mode runs on the host macOS machine. It opens available files in Microsoft Excel, Word, and PowerPoint, captures screenshots with `screencapture` when permission is available, and writes a GitHub-readable `INDEX.md` under:

```text
tools/manual-verification/evidence/v<version>-<git-sha>-macos/
```

## Intended Workflow

1. Generate or choose representative `.xlsx`, `.xlsm`, `.docx`, `.docm`, `.pptx`, and `.pptm` files.
2. Open each file in LibreOffice.
3. Confirm there is no repair dialog or visible corruption.
4. Save the file.
5. Reopen it.
6. Capture logs, screenshots, and a short summary.

Use the checklists in `checklists/` to keep verification consistent across releases.

## Boundaries

- Do not put GUI-dependent checks in `tests/DotnetPoi.*.Tests`.
- Do not commit large screenshots, videos, or generated Office files unless they are deliberately promoted to small fixtures.
- Keep Docker/GUI automation here, even when the command can be scripted.
