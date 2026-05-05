# DotnetPoi.Formula

**Formula evaluator for dotnet-poi — spreadsheet function evaluation as an optional add-on.**

`DotnetPoi.Formula` provides formula evaluation on top of `DotnetPoi.Core`. When this package is referenced, `createFormulaEvaluator()` automatically becomes available. Without it, the call throws `NotSupportedException`.

## Requirements

- [DotnetPoi.Core](../DotnetPoi.Core/README.md) — must be referenced (directly or transitively)

## When to use

| Use DotnetPoi.Formula if you need... | Don't need it if you only... |
|---|---|
| `IFormulaEvaluator.evaluate()` / `evaluateAll()` / `evaluateInCell()` | Read/write formula text and cached values |
| Evaluate SUM, AVERAGE, COUNT, MIN, MAX, CONCATENATE | Write/save/reload xlsx files |
| String concatenation with `&` | Read/write xls, docx, pptx, doc, ppt |

> All formula **text preservation** (`setCellFormula`, `getCellFormula`, cached `<v>` value) lives in Core and works without this package.

## How it works

`DotnetPoi.Formula` references `DotnetPoi.Core`, **not the other way around**. Registration happens automatically:

1. The static constructor of `FormulaEvaluator` calls `XSSFCreationHelper.RegisterFormulaEvaluatorFactory()`
2. When `ICreationHelper.createFormulaEvaluator()` is called and the factory is not yet registered, it uses lazy assembly discovery (`Type.GetType` + `RuntimeHelpers.RunClassConstructor`) to detect and initialize `DotnetPoi.Formula`
3. If found, the evaluator is created and returned
4. If not found, a clear `NotSupportedException` is thrown

No configuration, no manual registration — just add the NuGet reference.

## Usage example

```csharp
using DotnetPoi.SS.UserModel;
using DotnetPoi.XSSF.UserModel;

using var wb = new XSSFWorkbook();
var sheet = wb.createSheet();
var row = sheet.createRow(0);

row.createCell(0).setCellValue(10);
row.createCell(1).setCellValue(20);
row.createCell(2).setCellFormula("A1+B1");    // formula text — works without Formula package

// Evaluate — requires DotnetPoi.Formula
var evaluator = wb.getCreationHelper().createFormulaEvaluator();
evaluator.evaluateAll();

Console.WriteLine(row.getCell(2).getNumericCellValue()); // 30
```

## Supported formulas

### Arithmetic operators

| Operator | Description | Example |
|---|---|---|
| `+` | Addition | `A1+B1` |
| `-` | Subtraction | `A1-B1` |
| `*` | Multiplication | `A1*B1` |
| `/` | Division | `A1/B1` |
| `&` | String concatenation | `"hello "&"world"` |

### Cell references

| Syntax | Description | Example |
|---|---|---|
| `A1` | Single cell reference | `A1+10` |
| `A1:C1` | Range reference (in functions only) | `SUM(A1:C1)` |

### Functions

| Function | Signature | Description |
|---|---|---|
| `SUM` | `SUM(number1, [number2], ...)` | Sums numeric arguments. Non-numeric values are ignored. |
| `AVERAGE` | `AVERAGE(number1, [number2], ...)` | Arithmetic mean of numeric arguments. |
| `MIN` | `MIN(number1, [number2], ...)` | Minimum numeric value. |
| `MAX` | `MAX(number1, [number2], ...)` | Maximum numeric value. |
| `COUNT` | `COUNT(value1, [value2], ...)` | Counts numeric values in the argument list. |
| `CONCATENATE` | `CONCATENATE(text1, [text2], ...)` | Joins text arguments. Non-text values are converted to text. |

All functions accept range references (e.g. `SUM(A1:A5)`) and flatten nested ranges.

## Unsupported formulas

The following are **not yet implemented** and will throw `InvalidOperationException` with an appropriate message:

### Comparison and logical operators

| Operator | Expected behavior |
|---|---|
| `=` / `==` | Equality comparison |
| `<>` / `!=` | Inequality comparison |
| `<`, `>`, `<=`, `>=` | Relational comparison |
| `IF` | Conditional branching |

### Lookup and reference

| Function | Expected behavior |
|---|---|
| `VLOOKUP` | Vertical lookup |
| `HLOOKUP` | Horizontal lookup |
| `INDEX` | Index-based reference |
| `MATCH` | Position of value in range |
| `XLOOKUP` | Modern lookup (Excel 365) |
| `OFFSET` | Offset reference |
| `INDIRECT` | Indirect reference via text |

### Math and trigonometry

| Function | Expected behavior |
|---|---|
| `ABS` | Absolute value |
| `ROUND` | Round to specified digits |
| `ROUNDUP` / `ROUNDDOWN` | Round up/down |
| `INT` | Floor to integer |
| `MOD` | Modulo |
| `POWER` / `SQRT` | Exponentiation / square root |
| `CEILING` / `FLOOR` | Ceiling / floor |
| `RAND` / `RANDBETWEEN` | Random number generation |
| `PI` | π constant |
| `SIN` / `COS` / `TAN` | Trigonometric functions |
| `LOG` / `LN` | Logarithmic functions |

### Text

| Function | Expected behavior |
|---|---|
| `LEFT` / `RIGHT` / `MID` | String extraction |
| `LEN` | String length |
| `FIND` / `SEARCH` | String search |
| `REPLACE` / `SUBSTITUTE` | String replacement |
| `UPPER` / `LOWER` / `PROPER` | Case conversion |
| `TRIM` | Whitespace removal |
| `TEXT` | Number-to-text formatting |

### Date and time

| Function | Expected behavior |
|---|---|
| `TODAY` | Current date |
| `NOW` | Current date and time |
| `DATE` | Date from year/month/day |
| `YEAR` / `MONTH` / `DAY` | Date component extraction |
| `DATEDIF` | Date difference |
| `WEEKDAY` | Day of week |

### Logical

| Function | Expected behavior |
|---|---|
| `AND` | Logical AND |
| `OR` | Logical OR |
| `NOT` | Logical NOT |
| `IFERROR` | Error handling |
| `IFNA` | #N/A handling |
| `SWITCH` | Multi-condition branching |

### Statistical (advanced)

| Function | Expected behavior |
|---|---|
| `MEDIAN` | Median value |
| `MODE` | Most frequent value |
| `STDEV` / `STDEVP` | Standard deviation |
| `VAR` / `VARP` | Variance |
| `QUARTILE` / `PERCENTILE` | Quartile / percentile |
| `LARGE` / `SMALL` | K-th largest / smallest |
| `RANK` | Rank of value |
| `CORREL` | Correlation coefficient |

### Financial

| Function | Expected behavior |
|---|---|
| `PMT` | Loan payment |
| `FV` | Future value |
| `PV` | Present value |
| `NPV` / `IRR` | Net present value / internal rate of return |
| `RATE` | Interest rate |
| `NPER` | Number of periods |

### Information

| Function | Expected behavior |
|---|---|
| `ISBLANK` | Is cell blank |
| `ISNUMBER` / `ISTEXT` | Type checking |
| `ISERROR` / `ISERR` | Error checking |
| `TYPE` | Value type |
| `NA` | #N/A error |

### Rounding and precision

| Function | Expected behavior |
|---|---|
| `MROUND` | Round to nearest multiple |
| `TRUNC` | Truncate to integer |
| `EVEN` / `ODD` | Round to nearest even/odd |

## Version

This is a **v0.x** package. The supported function set will grow over time independently of `DotnetPoi.Core`, which targets v1.0 stability.

## License

[Apache License 2.0](../../LICENSE) — same as upstream Apache POI.
