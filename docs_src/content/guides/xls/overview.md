# xls (HSSF) Overview

HSSF provides read and write support for the legacy xls (BIFF) format. Coverage is ~10%.

## Status

HSSF is the oldest and least complete format implementation in dotnet-poi. It supports basic cell read/write operations but most advanced features (styles, images, charts, formulas, data validation, filters, pivot tables) are not implemented.

## Basic Read and Write

```csharp
using DotnetPoi.HSSF.UserModel;

// Write
using var wb = new HSSFWorkbook();
var sheet = wb.createSheet("Sheet1");
var row = sheet.createRow(0);
row.createCell(0).setCellValue("Hello from HSSF");

using var file = File.Create("output.xls");
wb.write(file);

// Read
using var stream = File.OpenRead("input.xls");
using var wb2 = new HSSFWorkbook(stream);
var value = wb2.getSheetAt(0).getRow(0).getCell(0).getStringCellValue();
```

## When to Use HSSF

- You need to read or update existing xls files without style changes
- You have a legacy workflow that requires xls output
- You are migrating from xls to xlsx and need a transitional solution

## When to Use XSSF Instead

- All new development should use XSSF (xlsx)
- xlsx is more capable, more widely supported, and has higher coverage
- Excel can open xlsx files from Excel 2007 onwards
- LibreOffice/OpenOffice handles xlsx as well as xls

## Limitations

- No formula support
- No image support
- No style support (fonts, fills, borders, number formats)
- No data validation
- No conditional formatting
- No auto filter
- No pivot tables
- No protection
- No chart support
- No rich text

HSSF is maintained for compatibility but is not a priority for new development.
