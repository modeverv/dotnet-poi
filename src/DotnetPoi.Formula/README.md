# DotnetPoi.Formula

[![CI](https://github.com/modeverv/dotnet-poi/actions/workflows/ci.yml/badge.svg)](https://github.com/modeverv/dotnet-poi/actions/workflows/ci.yml)
[![NuGet](https://img.shields.io/nuget/v/DotnetPoi.Formula)](https://www.nuget.org/packages/DotnetPoi.Formula)
[![NuGet Downloads](https://img.shields.io/nuget/dt/DotnetPoi.Formula)](https://www.nuget.org/packages/DotnetPoi.Formula)
![License](https://img.shields.io/badge/license-Apache%202.0-blue)
![.NET Standard](https://img.shields.io/badge/.NET%20Standard-2.0-%23820171)
![Status](https://img.shields.io/badge/status-early%20beta-yellow)

**Formula evaluator for dotnet-poi — evaluate spreadsheet formulas in xlsx workbooks.**

> ⚠️ This is an **unofficial** port and is **not affiliated with the Apache Software Foundation**. Apache POI is a registered trademark of the Apache Software Foundation.

---

## Install

```shell
dotnet add package DotnetPoi.Formula
```

```xml
<PackageReference Include="DotnetPoi.Formula" Version="..." />
```

> Requires **DotnetPoi.Core** 0.1.0+ and **.NET 8.0+** or **.NET Framework 4.7.2+**.

---

## Documentation

- [Getting started with formulas](https://modeverv.github.io/dotnet-poi/guides/xlsx/formulas)
- [Format coverage](https://modeverv.github.io/dotnet-poi/compatibility/format-coverage)

---

## When to use

| Use DotnetPoi.Formula if you need… | Don't need it if you only… |
|---|---|
| `IFormulaEvaluator.evaluate()` / `evaluateAll()` / `evaluateInCell()` | Read/write formula text and cached values |
| Programmatic access to freshly calculated results | Template fill → save → open in Excel |
| Use SUM, AVERAGE, COUNT, MIN, MAX, CONCATENATE | Write/save/reload xlsx files without calculation |

> All formula **text preservation** (`setCellFormula`, `getCellFormula`, cached `<v>` value) lives in **DotnetPoi.Core** and works without this package.

---

## How it works

`DotnetPoi.Formula` references `DotnetPoi.Core`, **not the other way around**. When this package is in your project, `createFormulaEvaluator()` is automatically enabled via lazy assembly discovery (`Type.GetType` + `RuntimeHelpers.RunClassConstructor`). No configuration needed — just add the NuGet reference.

---

## Usage

```csharp
using DotnetPoi.SS.UserModel;
using DotnetPoi.XSSF.UserModel;

using var wb = new XSSFWorkbook();
var sheet = wb.createSheet();
sheet.createRow(0).createCell(0).setCellValue(10);
sheet.getRow(0).createCell(1).setCellValue(20);
sheet.getRow(0).createCell(2).setCellFormula("A1+B1");

var evaluator = wb.getCreationHelper().createFormulaEvaluator();
evaluator.evaluateAll();

Console.WriteLine(sheet.getRow(0).getCell(2).getNumericCellValue()); // 30
```

---

## Supported formulas

### Operators

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
| `A1` | Single cell | `A1+10` |
| `A1:C1` | Range (in functions) | `SUM(A1:C1)` |

### Functions

| Function | Signature | Description |
|---|---|---|
| `SUM` | `SUM(number1, [number2], …)` | Sums numeric arguments. Non-numeric values ignored. |
| `AVERAGE` | `AVERAGE(number1, [number2], …)` | Arithmetic mean. |
| `MIN` | `MIN(number1, [number2], …)` | Minimum value. |
| `MAX` | `MAX(number1, [number2], …)` | Maximum value. |
| `COUNT` | `COUNT(value1, [value2], …)` | Counts numeric values. |
| `CONCATENATE` | `CONCATENATE(text1, [text2], …)` | Joins text arguments. |

---

## Unsupported (not yet implemented)

### Comparison / logical operators

`=`, `<>`, `<`, `>`, `<=`, `>=`, `IF`

### Lookup / reference

`VLOOKUP`, `HLOOKUP`, `INDEX`, `MATCH`, `XLOOKUP`, `OFFSET`, `INDIRECT`

### Math / trig

`ABS`, `ROUND`, `ROUNDUP`, `ROUNDDOWN`, `INT`, `MOD`, `POWER`, `SQRT`, `CEILING`, `FLOOR`, `RAND`, `RANDBETWEEN`, `PI`, `SIN`, `COS`, `TAN`, `LOG`, `LN`

### Text

`LEFT`, `RIGHT`, `MID`, `LEN`, `FIND`, `SEARCH`, `REPLACE`, `SUBSTITUTE`, `UPPER`, `LOWER`, `PROPER`, `TRIM`, `TEXT`

### Date / time

`TODAY`, `NOW`, `DATE`, `YEAR`, `MONTH`, `DAY`, `DATEDIF`, `WEEKDAY`

### Logical

`AND`, `OR`, `NOT`, `IFERROR`, `IFNA`, `SWITCH`

### Statistical

`MEDIAN`, `MODE`, `STDEV`, `STDEVP`, `VAR`, `VARP`, `QUARTILE`, `PERCENTILE`, `LARGE`, `SMALL`, `RANK`, `CORREL`

### Financial

`PMT`, `FV`, `PV`, `NPV`, `IRR`, `RATE`, `NPER`

### Information

`ISBLANK`, `ISNUMBER`, `ISTEXT`, `ISERROR`, `ISERR`, `TYPE`, `NA`

### Rounding

`MROUND`, `TRUNC`, `EVEN`, `ODD`

---

## Testing strategy

Accuracy of the formula engine is verified through:

- **Logic verification:** Unit tests for each supported function (`SUM`, `AVERAGE`, etc.) comparing results against known Excel outputs.
- **Reference handling:** Tests for relative/absolute cell references and range evaluation.
- **Interop calculation tests:** 
    - Verify that `DotnetPoi.Formula` produces the same results as Apache POI's `FormulaEvaluator` for a given set of inputs.
    - Verify that Excel recalculates the formulas to the same values when the file is opened.

---

## Version

This is a **v0.x** package. The supported function set will grow over time independently of `DotnetPoi.Core`, which targets v1.0 stability.

---

## Test coverage

| Project | Tests |
|---|---|
| Formula.Tests | 10 |

---

## License

[Apache License 2.0](../../LICENSE) — same as upstream Apache POI.

## Disclaimer

This project is not affiliated with the Apache Software Foundation or the Apache POI project.
Apache POI is a registered trademark of the Apache Software Foundation.
