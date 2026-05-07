# xls (HSSF) Overview

HSSF provides read and write support for the legacy xls (BIFF8) format. Coverage is ~35%.

## Status

HSSF has grown beyond the original bootstrap reader/writer. It now supports practical basic workbook operations: common cell types, multiple sheets, basic styles, layout records, representative Apache POI fixture loading, Java POI interop slices, and preservation of non-workbook OLE streams, VBA streams, and unknown BIFF records during light edits.

It is still not a full Apache POI HSSF port. Images, shapes, chart creation, comment editing, filters, pivots, and new BIFF formula token writing are not modeled through public APIs.

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
- You need to extract values from old xls files for indexing or migration
- You need light edits while preserving unmodeled OLE/BIFF content as much as possible

## When to Use XSSF Instead

- All new development should use XSSF (xlsx)
- xlsx is more capable, more widely supported, and has higher coverage
- Excel can open xlsx files from Excel 2007 onwards
- LibreOffice/OpenOffice handles xlsx as well as xls

## Supported Today

- String, numeric, boolean, blank, and error cells
- Multiple sheets, sparse rows/cells, and high column indexes
- Basic font and cell style round-trips: font name/size/bold/italic/color, data formats, alignment, wrap, borders, fills
- Column width, row height, hidden rows/columns, merged regions, freeze panes
- Reading formula text and cached formula values from existing files
- Representative Apache POI `.xls` fixture loading
- Java POI interop fixtures for basic, styles, layout, unicode, and comprehensive workbooks
- Preservation of non-Workbook OLE streams, macro/VBA streams, directory metadata, and unknown BIFF records

## Limitations

- No formula evaluation
- No new BIFF formula token writing
- No image or drawing usermodel
- No chart creation/editing
- No comment or hyperlink editing API
- No data validation
- No auto filter
- No pivot tables
- No rich text

HSSF is maintained for compatibility but is not a priority for new development.
