# Pivot Tables

dotnet-poi supports programmatic creation of pivot tables. Editing existing pivot tables is not modeled but they are preserved on round-trip.

## Creating a Pivot Table

```csharp
using DotnetPoi.XSSF.UserModel;
using DotnetPoi.SS.Util;

// Create source data
var sheet = wb.createSheet("Data");
sheet.createRow(0).CreateCells(new[] { "Category", "Amount" });
sheet.createRow(1).CreateCells(new[] { "Food", "100" });
sheet.createRow(2).CreateCells(new[] { "Travel", "200" });

// Create pivot table sheet
var pivotSheet = wb.createSheet("Pivot");

// Define source range
var sourceRange = new CellRangeAddress(0, 2, 0, 1);  // Data!A1:B3

// Create pivot table
var pivotTable = pivotSheet.createPivotTable(sourceRange, new CellReference(0, 0));

// Add rows and data fields
var categoryField = pivotTable.getRowLabel().AddValueField(0);     // Column 0 = Category
var amountField = pivotTable.getDataFields().AddValueField(1);     // Column 1 = Amount
amountField.SetValueFieldSum();
```

## Limitations

- Only programmatic creation is supported
- Editing existing pivot tables is not modeled
- Pivot tables in existing files are preserved as unknown parts on round-trip
