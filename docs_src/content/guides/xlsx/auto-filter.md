# Auto Filter

Add dropdown filter buttons to column headers.

## Basic Usage

```csharp
sheet.setAutoFilter(new CellRangeAddress(2, 6, 0, 3));
// Adds filter buttons to the header range A2:D6
```

The method takes a `CellRangeAddress` specifying the data range including headers. Excel displays dropdown arrows in the first row of the range.

## Reading Auto Filter

```csharp
var filter = sheet.getAutoFilter();
if (filter is not null)
{
    var range = filter.CellRange;
    // range.FirstRow, range.LastRow, range.FirstCol, range.LastCol
}
```

Auto filter settings are preserved on round-trip.
