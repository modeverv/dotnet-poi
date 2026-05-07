# Your First Spreadsheet

This page walks through creating an `.xlsx` file with string and numeric cells.

## Minimal Example

```csharp
using DotnetPoi.XSSF.UserModel;

using var workbook = new XSSFWorkbook();
var sheet = workbook.createSheet("Sheet1");

var row = sheet.createRow(0);
row.createCell(0).setCellValue("Hello");
row.createCell(1).setCellValue(42);

using var file = File.Create("output.xlsx");
workbook.write(file);
```

## Step by Step

### 1. Create a Workbook

```csharp
using var workbook = new XSSFWorkbook();
```

`XSSFWorkbook` represents an entire Excel workbook (the `.xlsx` package). Dispose it when done to free resources.

### 2. Create a Sheet

```csharp
var sheet = workbook.createSheet("Sheet1");
```

Sheets are named tabs in the workbook. The name must be unique within the workbook.

### 3. Create Rows and Cells

```csharp
var row = sheet.createRow(0);        // first row (0-indexed)
row.createCell(0).setCellValue("A string");
row.createCell(1).setCellValue(123.45);
row.createCell(2).setCellValue(42);   // int is stored as double
```

Cells are created at a column index (0-indexed). Uncreated cells are blank when the file is opened.

### 4. Save

```csharp
using var file = File.Create("output.xlsx");
workbook.write(file);
```

The `write()` method serializes the entire workbook to the stream. The stream is a standard .NET `Stream`, so you can write to a file, a `MemoryStream`, or an HTTP response body.

### 5. Read It Back

```csharp
using var stream = File.OpenRead("output.xlsx");
using var workbook = new XSSFWorkbook(stream);

var sheet = workbook.getSheetAt(0);
var row = sheet.getRow(0);

Console.WriteLine(row.getCell(0).getStringCellValue()); // "A string"
Console.WriteLine(row.getCell(1).getNumericCellValue()); // 123.45
```

## Full Runnable Example

See `examples/Phase0WriteExample/` in the repository for a complete example that writes an xlsx, saves it, and prints confirmation:

[examples/Phase0WriteExample](https://github.com/modeverv/dotnet-poi/tree/master/examples/Phase0WriteExample)

Run it:

```bash
dotnet run --project examples/Phase0WriteExample
```

Generated file: `examples/output/phase0-write-example.xlsx`

## What's Next

- [Cell Types and Values](../guides/xlsx/cell-types.md) — strings, numbers, dates, booleans, formulas
- [Styles and Formatting](../guides/xlsx/styles.md) — fonts, fills, borders, number formats
