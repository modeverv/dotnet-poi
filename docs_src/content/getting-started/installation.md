# Installation

dotnet-poi ships as **two separate NuGet packages**. Most projects only need `DotnetPoi.Core`.

## Package Reference

Add the package to your `.csproj`:

```xml
<ItemGroup>
  <PackageReference Include="DotnetPoi.Core" Version="0.1.0" />
</ItemGroup>
```

If you need spreadsheet formula evaluation, add the Formula package as well:

```xml
<ItemGroup>
  <PackageReference Include="DotnetPoi.Core" Version="0.1.0" />
  <PackageReference Include="DotnetPoi.Formula" Version="0.1.0" />
</ItemGroup>
```

## Why Two Packages?

| Package | Contents | When to use |
|---|---|---|
| `DotnetPoi.Core` | All format implementations (xlsx, xls, docx, pptx, doc, ppt) + common interfaces + XML writer | Always required |
| `DotnetPoi.Formula` | Formula evaluator (`IFormulaEvaluator`, `FormulaEvaluator`, `CellValue`) | Only when you need programmatic formula evaluation |

**Design principle:** `Core` has zero knowledge of `Formula`. Adding `DotnetPoi.Formula` to your project automatically enables `createFormulaEvaluator()` via lazy assembly discovery at runtime. Without it, the call throws a clear `NotSupportedException`.

## Verify Installation

Create a file `TestInstall.csproj`:

```xml
<Project Sdk="Microsoft.NET.Sdk">
  <PropertyGroup>
    <OutputType>Exe</OutputType>
    <TargetFramework>net8.0</TargetFramework>
  </PropertyGroup>
  <ItemGroup>
    <PackageReference Include="DotnetPoi.Core" Version="0.1.0" />
  </ItemGroup>
</Project>
```

Add `Program.cs`:

```csharp
using DotnetPoi.XSSF.UserModel;

using var wb = new XSSFWorkbook();
var sheet = wb.createSheet("Test");
var row = sheet.createRow(0);
row.createCell(0).setCellValue("Hello, dotnet-poi!");

using var file = File.Create("test.xlsx");
wb.write(file);

Console.WriteLine("test.xlsx created successfully.");
```

Run:

```bash
dotnet run
```

Open `test.xlsx` in Excel or LibreOffice to confirm the content.

## Package Split Details

See the [Package Split](../compatibility/package-split.md) page for the full design rationale.
