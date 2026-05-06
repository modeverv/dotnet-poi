# Conditional Formatting

Apply formatting rules that change based on cell values.

## Basic Example

```csharp
using DotnetPoi.XSSF.UserModel;

var conditional = new XSSFConditionalFormatting { Sqref = "D4:D6" };
conditional.Rules.Add(new XSSFCFRule
{
    Type = ConditionalFormatType.CellIs,
    Operator = "greaterThan",
    Priority = 1,
    DxfId = -1
});
conditional.Rules[0].Formulas.Add("20");

sheet.AddConditionalFormatting(conditional);
```

This highlights cells in range D4:D6 where the value is greater than 20.

## Conditional Format Types

| Type | Description |
|---|---|
| `CellIs` | Compare cell value against a formula |
| `Formula` | Custom formula expression |
| `ColorScale` | Color gradient (2 or 3 colors) |
| `DataBar` | Data bar |
| `IconSet` | Icon set |

## Full Runnable Example

See `examples/UsageSamples/Program.cs` (`CreateSpreadsheet`):

[examples/UsageSamples](https://github.com/modeverv/dotnet-poi/tree/main/examples/UsageSamples)
