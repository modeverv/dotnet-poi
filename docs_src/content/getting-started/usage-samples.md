# Usage Samples

This documentation starts with runnable examples, not API catalogs. The first sample project creates real Office files, saves them under `examples/output/`, then reads them back to verify the generated content.

Run it from the repository root:

```bash
dotnet run --project examples/UsageSamples/UsageSamples.csproj
```

Generated files:

```text
examples/output/usage-workbook.xlsx
examples/output/usage-document.docx
examples/output/usage-presentation.pptx
```

## Spreadsheet

`usage-workbook.xlsx` demonstrates the spreadsheet tasks most users need first:

- create an `.xlsx` workbook and sheet
- write string and numeric cells
- apply fonts, fills, borders, alignment, and number formats
- merge a title range
- freeze header rows
- add a whole-number data validation
- preserve rich text in a shared string
- read the workbook back and assert cell values

Source: `examples/UsageSamples/Program.cs`, `CreateSpreadsheet`.

## Word Document

`usage-document.docx` demonstrates:

- create paragraphs and runs
- apply bold text, font size, color, underline, and hyperlink metadata
- create a table
- embed an inline JPEG image
- read the document back and assert paragraphs, table content, and image data

Source: `examples/UsageSamples/Program.cs`, `CreateDocument`.

## PowerPoint Presentation

`usage-presentation.pptx` demonstrates:

- create slides
- add text boxes and formatted runs
- place and rotate an image
- create a simple table on a slide
- read the presentation back and assert slide, text, and table content

Source: `examples/UsageSamples/Program.cs`, `CreatePresentation`.

## Documentation Scope

Phase 9 documentation is usage-only for now. Do not add class hierarchy pages or generated API reference pages until the practical guides have enough runnable coverage.
