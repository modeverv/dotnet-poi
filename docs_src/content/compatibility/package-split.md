# Package Split: Core vs Formula

dotnet-poi ships as **two separate NuGet packages** to decouple Core stability from Formula maturity.

## The Two Packages

| Package | NuGet ID | Contents | Stability |
|---|---|---|---|
| Core | `DotnetPoi.Core` | All format implementations (XSSF/xlsx, HSSF/xls, XWPF/docx, XSLF/pptx, HWPF/doc, HSLF/ppt, POIFS) + common interfaces + XML writer | **Stable** — ready for v1.0 |
| Formula | `DotnetPoi.Formula` | Limited formula evaluator (`IFormulaEvaluator`, `FormulaEvaluator`, `CellValue`) | **Experimental** — intentionally small subset |

## Design Rules

1. **Core has zero knowledge of Formula.** No compile-time dependency. Core works perfectly without Formula.
2. **Formula references Core.** It consumes Core types like `ICell`, `IWorkbook`.
3. **Auto-discovery at runtime.** Formula registers itself via a static constructor. When you call `createFormulaEvaluator()`, Core uses lazy assembly discovery (`Type.GetType` + `RuntimeHelpers.RunClassConstructor`) to find and activate Formula automatically.
4. **Graceful fallback.** If Formula is not installed, `createFormulaEvaluator()` throws a clear `NotSupportedException`.

## What Goes Where

| Feature | Package | Reason |
|---|---|---|
| Read/write xlsx, xls, docx, pptx | Core | Core format operations |
| Cell values, styles, layout | Core | No evaluation needed |
| Formula text read/write | Core | Preserving formula text + cached values |
| `setCellFormula` / `getCellFormula` | Core | Text operations |
| `calcPr fullCalcOnLoad` | Core | Tell Excel to recalculate on open |
| `createFormulaEvaluator()` | Core (factory), Formula (impl) | Factory in Core, evaluator in Formula |
| `evaluate()` / `evaluateAll()` for supported simple expressions | Formula | Limited calculation subset |
| `evaluateFormulaCell()` for supported simple expressions | Formula | Refresh cached result when the formula is in scope |

## Why This Split

- **Core can be stable early.** All spreadsheet read/write/format logic is self-contained.
- **Formula can stay narrow.** Evaluation is complex; full Excel-compatible calculation is not a current project goal.
- **Smaller dependency for simple use cases.** Users who only need xlsx read/write don't pull in the formula engine.
- **Security.** Applications handling untrusted documents can omit the formula evaluator entirely, reducing attack surface.

## Installation

```xml
<!-- Read/write only -->
<PackageReference Include="DotnetPoi.Core" Version="..." />

<!-- Add formula evaluation -->
<PackageReference Include="DotnetPoi.Core" Version="..." />
<PackageReference Include="DotnetPoi.Formula" Version="..." />
```
