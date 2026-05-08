# Formulas

dotnet-poi supports writing formula text and preserving cached values. Full formula evaluation is not available.

## Writing a Formula

```csharp
var cell = row.createCell(0);
cell.setCellFormula("SUM(A1:A10)");
cell.setCellValue(55.0);  // optional cached result
```

When you set a formula and save, Excel will recalculate the formula when the file is opened. Setting the cached value is optional but recommended for interop.

## Reading a Formula

```csharp
var formula = cell.getCellFormula();                    // "SUM(A1:A10)"
var cachedType = cell.getCachedFormulaResultType();     // CellType.Numeric
var cachedValue = cell.getNumericCellValue();           // 55.0
```

## Tell Excel to Recalculate on Open

```csharp
workbook.getCTWorkbook().calcPr.fullCalcOnLoad = true;
```

This sets the `fullCalcOnLoad` flag in the workbook. When the file is opened in Excel, all formulas are recalculated regardless of cached values.

## Formula Evaluation

Full formula evaluation is **permanently deferred**. The `DotnetPoi.Formula` package contains a limited evaluator for simple arithmetic and a small function subset such as SUM, AVERAGE, COUNT, MIN, MAX, and CONCATENATE. It is not an Excel-compatible calculation engine.

Without `DotnetPoi.Formula`, calling `createFormulaEvaluator()` throws `NotSupportedException`.

See [Package Split](../../compatibility/package-split.md) for details.

## Full Runnable Example

See `examples/Phase5FormulaEvaluatorExample/`:

[examples/Phase5FormulaEvaluatorExample](https://github.com/modeverv/dotnet-poi/tree/master/examples/Phase5FormulaEvaluatorExample)
