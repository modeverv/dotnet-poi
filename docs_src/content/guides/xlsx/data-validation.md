# Data Validation

Add input rules to cells to restrict what values users can enter.

## Whole Number Validation

```csharp
using DotnetPoi.XSSF.UserModel;
using DotnetPoi.SS.UserModel;

sheet.AddDataValidation(new XSSFDataValidation
{
    Sqref = "C4:C6",              // cell range to validate
    Type = DataValidationType.Whole,
    Operator = DataValidationOperator.Between,
    Formula1 = "1",
    Formula2 = "100",
    PromptTitle = "Quantity",
    PromptMessage = "Enter a whole number from 1 to 100.",
    ErrorTitle = "Invalid quantity",
    ErrorMessage = "Quantity must be between 1 and 100."
});
```

## Validation Types

| Type | Description |
|---|---|
| `Whole` | Integer numbers |
| `Decimal` | Decimal numbers |
| `List` | Value must be in a list |
| `Date` | Date values |
| `Time` | Time values |
| `TextLength` | Text character count |
| `Custom` | Custom formula |

## Operators

`Between`, `NotBetween`, `Equal`, `NotEqual`, `GreaterThan`, `LessThan`, `GreaterOrEqual`, `LessOrEqual`.

## Full Runnable Example

See `examples/UsageSamples/Program.cs` (`CreateSpreadsheet`):

[examples/UsageSamples](https://github.com/modeverv/dotnet-poi/tree/master/examples/UsageSamples)
