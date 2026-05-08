# Package Split Architecture

dotnet-poi uses a **multi-package architecture** with clear separation of concerns. This was designed to let OOXML stability advance independently of legacy binary format development.

## The Packages

| Package | NuGet ID | Contents | Stability |
|---|---|---|---|
| Ooxml | `DotnetPoi.Ooxml` | XSSF (xlsx/xlsm), XWPF (docx/docm), XSLF (pptx/pptm) + OPC/openxml + POIFS + Common | **Stable** — ready for v1.0 |
| Legacy | `DotnetPoi.Legacy` | HSSF (xls), HWPF (doc), HSLF (ppt) + POIFS + Common | **In development** |
| Formula | `DotnetPoi.Formula` | Limited formula evaluator (`IFormulaEvaluator`, `FormulaEvaluator`, `CellValue`) | **Experimental** — intentionally small subset |
| Common | `DotnetPoi.Common` | SS interfaces, shared enums, common exceptions, XML writer foundation | **Stable** (transitive) |
| POIFS | `DotnetPoi.POIFS` | OLE2/CFB compound file container, encryption helpers | **Stable** (transitive) |
| **All** | `DotnetPoi.All` | Meta-package referencing all of the above | **Smoke-tested** |

## Design Rules

1. **Format packages have zero knowledge of `Formula`.** No compile-time dependency. Ooxml and Legacy work perfectly without Formula.
2. **Formula references Ooxml (or Legacy).** It consumes types like `ICell`, `IWorkbook`.
3. **Auto-discovery at runtime.** Formula registers itself via a static constructor. When you call `createFormulaEvaluator()`, the format package uses lazy assembly discovery (`Type.GetType` + `RuntimeHelpers.RunClassConstructor`) to find and activate Formula automatically.
4. **Graceful fallback.** If Formula is not installed, `createFormulaEvaluator()` throws a clear `NotSupportedException`.

## What Goes Where

| Feature | Package | Reason |
|---|---|---|
| Read/write xlsx, docx, pptx | Ooxml | OOXML format operations |
| Read/write xls, doc, ppt | Legacy | Legacy binary format operations |
| Cell values, styles, layout | Ooxml (xlsx) / Legacy (xls) | No evaluation needed |
| Formula text read/write | Ooxml (xlsx) / Legacy (xls) | Preserving formula text + cached values |
| `setCellFormula` / `getCellFormula` | Ooxml (xlsx) / Legacy (xls) | Text operations |
| `calcPr fullCalcOnLoad` | Ooxml | Tell Excel to recalculate on open |
| `createFormulaEvaluator()` | Ooxml/Legacy (factory), Formula (impl) | Factory in format package, evaluator in Formula |
| `evaluate()` / `evaluateAll()` for supported simple expressions | Formula | Limited calculation subset |
| `evaluateFormulaCell()` for supported simple expressions | Formula | Refresh cached result when the formula is in scope |

## Why This Split

- **Ooxml can be stable early.** All OOXML read/write/format logic is self-contained. No dependency on legacy binary development.
- **Legacy can evolve safely.** HSSF/HWPF/HSLF development can proceed without destabilizing OOXML users.
- **Formula can stay narrow.** Evaluation is complex; full Excel-compatible calculation is not a current project goal.
- **Smaller dependency for simple use cases.** Users who only need xlsx read/write don't pull in legacy format code.
- **Security.** Applications handling untrusted documents can omit the formula evaluator entirely, reducing attack surface.

## Installation Examples

```xml
<!-- OOXML only: xlsx / docx / pptx -->
<PackageReference Include="DotnetPoi.Ooxml" Version="..." />

<!-- OOXML + formula evaluation -->
<PackageReference Include="DotnetPoi.Ooxml" Version="..." />
<PackageReference Include="DotnetPoi.Formula" Version="..." />

<!-- Legacy binary only: xls / doc / ppt -->
<PackageReference Include="DotnetPoi.Legacy" Version="..." />

<!-- Legacy + formula evaluation -->
<PackageReference Include="DotnetPoi.Legacy" Version="..." />
<PackageReference Include="DotnetPoi.Formula" Version="..." />

<!-- Everything with one dependency -->
<PackageReference Include="DotnetPoi.All" Version="..." />
```
