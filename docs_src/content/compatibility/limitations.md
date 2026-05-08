# Limitations

This page documents known limitations and missing features that may affect real-world use. Check [Format Coverage](format-coverage.md) for the full status matrix.

## Critical Gaps

### Formula Evaluation

dotnet-poi can write formula text and preserve cached values, but **full formula evaluation is permanently deferred**. This means:

- `setCellFormula("SUM(A1:A10)")` works and Excel recalculates on open
- `XSSFFormulaEvaluator` exists as a partial implementation (SUM, AVERAGE, CONCATENATE, basic arithmetic) but will not be expanded
- Programmatic access to freshly calculated results without saving is **not supported**
- Adding `DotnetPoi.Formula` package enables `createFormulaEvaluator()` for the limited evaluator

### Chart Creation

Charts in xlsx and pptx are **preserved on round-trip but cannot be created or edited programmatically**. If your workflow requires generating chart visuals from data, this is not yet supported.

### docx Style Editing

Paragraph style references are supported with `setStyle()` / `getStyleID()`, and `word/styles.xml` is read or generated for new documents. Character styles, table styles, and full Word style inheritance are not yet editable or evaluated by dotnet-poi.

### docx Table Styling

docx tables support basic creation, width, grid columns, row height, header rows, cell width, horizontal grid span, vertical merge, and vertical alignment. Detailed border and shading creation is not yet exposed as an API, but existing `tblPr` / `trPr` / `tcPr` children are preserved as raw XML during round-trip.

## Moderate Gaps

| Feature | Format | Impact |
|---|---|---|
| Comments (read/create) | xlsx, docx | Existing comments preserved on round-trip but cannot be read or created |
| Content controls (SDT) | docx | Block-level and inline SDT are preserved on round-trip but cannot be edited through a public API |

| Track changes | docx | Revision marks not supported |
| Grouped shapes | pptx | Preserved as raw `spTree` XML but cannot be edited through a public API |
| Notes slides | pptx | Existing notes slides are preserved but cannot be created or read through a public API |
| Sparklines | xlsx | In-cell sparkline charts not supported |
| External data connections | xlsx | Existing connection parts are preserved but cannot be edited through a public API |
| HSSF advanced object model | xls | Basic workbook values, styles/layout slices, interop, and preservation exist, but images, charts, comment editing, filters, pivots, and new formula writing are not modeled |
| HWPF advanced object model | doc | Body text extraction and limited body edits work, but tables, images, header/footer stories, footnotes, comments, and fields are not modeled through public APIs |

## Minor Gaps

| Feature | Format | Notes |
|---|---|---|
| Auto-shapes | xlsx | Preserved as unknown DrawingML XML; only pictures are modeled |
| SmartArt editing | pptx | Preserved as unknown parts |
| Animations/transitions editing | pptx | Preserved as unknown parts |
| Theme editing | pptx | Layout/master/theme preserved but not editable |
| Password hashing | xlsx | Protection on/off works; password hash not implemented |
| ppt (HSLF) | ppt | Minimal reader only; no no-op write/interoperability track completed yet |

## When to Use Java POI Instead

If your project depends on any of the following, consider using Apache POI (Java) directly or bridging via a microservice:

- Programmatic formula evaluation (Excel calculation engine)
- Chart creation from data
- Editing character/table styles or resolving complex Word style inheritance
- Full xls (BIFF) support beyond the current basic values/styles/layout/preservation slices
- Legacy doc (HWPF) features beyond body text extraction and limited body edits
- Legacy ppt (HSLF) workflows beyond the minimal reader
