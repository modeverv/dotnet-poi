# Auto Filter

Add dropdown filter buttons to column headers.

## Basic Usage

```csharp
sheet.setAutoFilter(new CellRangeAddress(2, 6, 0, 3));
// Adds filter buttons to the header range A3:D7
```

The method takes a `CellRangeAddress` specifying the data range including headers. Excel displays dropdown arrows in the first row of the range.

## Reading Auto Filter

```csharp
var filter = sheet.getAutoFilter();
if (filter is not null)
{
    Console.WriteLine(filter.FormatAsString());
    // filter.FirstRow, filter.LastRow, filter.FirstCol, filter.LastCol
}
```

Auto filter settings are preserved on round-trip.

## Full Runnable Example

See `examples/UsageSamples/Program.cs` (`CreateSpreadsheet`):

[examples/UsageSamples](https://github.com/modeverv/dotnet-poi/tree/master/examples/UsageSamples)
