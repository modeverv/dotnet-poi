# Layout

Column widths, row heights, merged regions, freeze panes, hidden rows/columns, and active cell position.

## Column Width

```csharp
sheet.setColumnWidth(0, 24 * 256);  // column A = 24 characters wide
```

Width is in units of 1/256th of a character width. Multiply by 256 for convenience.

To read back the width of a column:

```csharp
int width = sheet.getColumnWidth(0);  // returns width in 1/256th character units
```

## Row Height

```csharp
row.setHeight(30);  // 30 points
```

Height is in points (1/72 inch). The default row height is 15.0 points.

To read back the height of a row:

```csharp
float height = row.getHeight();  // returns height in points
```

## Merged Regions

```csharp
using DotnetPoi.SS.Util;

sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 3));
// Merges cells A1:D1 (row 0, columns 0-3)
```

`CellRangeAddress` constructor: `(firstRow, lastRow, firstCol, lastCol)`.

## Freeze Panes

```csharp
sheet.createFreezePane(0, 3);
// Freezes the first 3 rows (header area) — scrollable below row 3

sheet.createFreezePane(2, 0);
// Freezes the first 2 columns — scrollable to the right
```

Parameters: `(colSplit, rowSplit)`. The split is 0-indexed; rows above and columns left of the split stay visible.

## Hidden Rows and Columns

```csharp
sheet.getRow(2).setZeroHeight(true);  // hide row 3

sheet.setColumnHidden(1, true);       // hide column B
sheet.isColumnHidden(1);              // check if hidden
```

Hidden rows/columns are preserved on round-trip.

## Active Cell and Sheet Selection

In Japanese spreadsheet workflows, setting the active cell position and selecting the active sheet is particularly important for user navigation when the file is opened.

### Set the Active Cell

```csharp
sheet.setActiveCell("D5");  // cell D5 will be selected when opened
```

The active cell is the cell that has keyboard focus when the workbook is opened in Excel. The reference is specified as an Excel-style string (e.g. "A1", "C10", "AB42").

To read back the active cell:

```csharp
string? activeCell = sheet.getActiveCell();  // returns "D5", or null if not set
```

### Set the Active Sheet

```csharp
wb.setActiveSheet(1);   // the second sheet (0-indexed) becomes the active tab
wb.setSelectedTab(1);   // alias — selects the tab
```

To read back which sheet is active:

```csharp
int activeSheet = wb.getActiveSheetIndex();  // returns 0, 1, 2, ...
```

### Select a Sheet

When a workbook has multiple sheets, you can mark a sheet as selected (shown in the tab bar):

```csharp
sheet.setSelected(true);   // mark this sheet as selected
bool selected = sheet.isSelected();  // check if selected
```

### Notes

- `setActiveSheet()` and `getActiveSheetIndex()` are serialized to XML and round-trip correctly.
- `setActiveCell()` and `sheet.setSelected()` are available in the API but are currently **in-memory only** — they are not yet written to or read from the underlying XML. A future update will add XML serialization for the `<selection>` element and `activeCell` attribute in `<sheetViews>`.

See `examples/UsageSamples/Program.cs` (`CreateSpreadsheet`) and `examples/Phase3InterfaceExample/`:

[examples/UsageSamples](https://github.com/modeverv/dotnet-poi/tree/master/examples/UsageSamples)

[examples/Phase3InterfaceExample](https://github.com/modeverv/dotnet-poi/tree/master/examples/Phase3InterfaceExample)
