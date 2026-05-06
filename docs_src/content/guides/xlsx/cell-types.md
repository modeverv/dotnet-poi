# Cell Types and Values

dotnet-poi supports all standard cell types: string, numeric, boolean, date, formula (text + cached value), and error.

## Writing Cell Values

```csharp
using DotnetPoi.XSSF.UserModel;

using var wb = new XSSFWorkbook();
var ws = wb.createSheet("Types");
var row = ws.createRow(0);

row.createCell(0).setCellValue("Hello");       // String
row.createCell(1).setCellValue(42.5);           // Numeric (double)
row.createCell(2).setCellValue(true);           // Boolean
row.createCell(3).setCellValue(42);             // int is stored as double
```

## Reading Cell Values

```csharp
using var stream = File.OpenRead("input.xlsx");
using var wb = new XSSFWorkbook(stream);
var ws = wb.getSheetAt(0);
var row = ws.getRow(0);

var cellType = row.getCell(0).getCellType();  // CellType enum

switch (cellType)
{
    case CellType.Numeric:
        var num = row.getCell(0).getNumericCellValue();
        break;
    case CellType.String:
        var str = row.getCell(0).getStringCellValue();
        break;
    case CellType.Boolean:
        var b = row.getCell(0).getBooleanCellValue();
        break;
    case CellType.Formula:
        var formula = row.getCell(0).getCellFormula();
        var cachedType = row.getCell(0).getCachedFormulaResultType();
        break;
    case CellType.Error:
        var err = row.getCell(0).getErrorCellString();
        break;
    case CellType.Blank:
        // Cell is empty
        break;
}
```

## Date Values

Dates in xlsx are stored as numeric cells with a date format. Write a `DateTime` using `setCellValue`:

```csharp
row.createCell(0).setCellValue(new DateTime(2025, 1, 15));
```

To display as a date, apply a date format:

```csharp
var style = wb.createCellStyle();
style.setDataFormat(wb.createDataFormat().getFormat("yyyy-MM-dd"));
cell.setCellStyle(style);
```

When reading back, use `DateUtil.isCellDateFormatted()` to check if a numeric cell is a date:

```csharp
if (DateUtil.isCellDateFormatted(cell))
{
    var date = cell.getDateCellValue();  // returns DateTime
}
```

## Formula Cells

Write a formula with `setCellFormula`. The cached result can be set separately:

```csharp
var cell = row.createCell(0);
cell.setCellFormula("SUM(A1:A10)");
cell.setCellValue(55.0);  // cached result — Excel recalculates on open
```

Read formula text and cached value:

```csharp
var formula = cell.getCellFormula();                 // "SUM(A1:A10)"
var cachedType = cell.getCachedFormulaResultType();  // CellType.Numeric
var cachedValue = cell.getNumericCellValue();        // 55.0
```

**Note:** Full formula evaluation is not available. See [Package Split](../../compatibility/package-split.md).

## Full Runnable Example

See `examples/Phase7CellTypesExample/`:

[examples/Phase7CellTypesExample](https://github.com/modeverv/dotnet-poi/tree/main/examples/Phase7CellTypesExample)
