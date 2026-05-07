# Pivot Tables

dotnet-poi supports programmatic creation of pivot tables. Editing existing pivot tables is not modeled but they are preserved on round-trip.

## Creating a Pivot Table

```csharp
using DotnetPoi.XSSF.UserModel;

// Create source data
var sheet = wb.createSheet("Data");
var header = sheet.createRow(0);
header.createCell(0).setCellValue("Category");
header.createCell(1).setCellValue("Amount");
var food = sheet.createRow(1);
food.createCell(0).setCellValue("Food");
food.createCell(1).setCellValue(100);

// Create pivot table sheet
var pivotSheet = wb.createSheet("Pivot");

// Create pivot table
var pivotTable = pivotSheet.createPivotTable("A1", "A1:B2", "Data");

// Add rows and data fields
pivotTable.RowLabels.Add(0);   // Column 0 = Category
pivotTable.DataColumns.Add(1); // Column 1 = Amount
```

## Limitations

- Only programmatic creation is supported
- Editing existing pivot tables is not modeled
- Pivot tables in existing files are preserved as unknown parts on round-trip

## Full Runnable Example

See `examples/UsageSamples/Program.cs` (`CreateSpreadsheet`):

[examples/UsageSamples](https://github.com/modeverv/dotnet-poi/tree/main/examples/UsageSamples)
