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

### docx Styles

Only direct formatting (bold, italic, font size, color, etc.) is supported. **Paragraph styles (Normal, Heading 1, Title) and character styles are not implemented.** Documents that depend on style-based formatting may lose appearance when round-tripped.

### docx Table Features

Cell merging and table borders are **not implemented** in docx tables.

## Moderate Gaps

| Feature | Format | Impact |
|---|---|---|
| Comments (read/create) | xlsx, docx | Existing comments preserved on round-trip but cannot be read or created |
| Text boxes (w:txbxContent) | docx | Text inside Word text boxes is not readable |
| Content controls (SDT) | docx | Structured document tags not supported |
| Section columns | docx | Multi-column layout not supported |
| Track changes | docx | Revision marks not supported |
| Grouped shapes | pptx | Shape grouping not supported |
| Notes slides | pptx | Speaker notes cannot be created or read |
| Sparklines | xlsx | In-cell sparkline charts not supported |
| External data connections | xlsx | Data connections not supported |

## Minor Gaps

| Feature | Format | Notes |
|---|---|---|
| Auto-shapes | xlsx | Not supported |
| SmartArt editing | pptx | Preserved as unknown parts |
| Animations/transitions editing | pptx | Preserved as unknown parts |
| Theme editing | pptx | Layout/master/theme preserved but not editable |
| Password hashing | xlsx | Protection on/off works; password hash not implemented |
| xls (HSSF) | xls | Legacy format; minimal support |

## When to Use Java POI Instead

If your project depends on any of the following, consider using Apache POI (Java) directly or bridging via a microservice:

- Programmatic formula evaluation (Excel calculation engine)
- Chart creation from data
- Heavy use of docx paragraph/table styles
- Full xls (BIFF) support
- Legacy doc (HWPF) or ppt (HSLF) formats
