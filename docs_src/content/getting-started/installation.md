# Installation

dotnet-poi ships as **multiple NuGet packages** — pick the combination that fits your project.

## Quick Pick

| If you need… | Install this |
|---|---|
| xlsx / docx / pptx read/write only | `DotnetPoi.Ooxml` |
| OOXML + formula evaluator | `DotnetPoi.Ooxml` + `DotnetPoi.Formula` |
| xls / doc / ppt (legacy binary) only | `DotnetPoi.Legacy` |
| All formats + formula (everything) | `DotnetPoi.All` |

## Package Reference

**OOXML-only project (xlsx, docx, pptx):**

```xml
<ItemGroup>
  <PackageReference Include="DotnetPoi.Ooxml" Version="..." />
</ItemGroup>
```

**Add formula evaluation when needed:**

```xml
<ItemGroup>
  <PackageReference Include="DotnetPoi.Ooxml" Version="..." />
  <PackageReference Include="DotnetPoi.Formula" Version="..." />
</ItemGroup>
```

**Legacy binary formats (xls, doc, ppt):**

```xml
<ItemGroup>
  <PackageReference Include="DotnetPoi.Legacy" Version="..." />
</ItemGroup>
```

**Everything in one dependency:**

```xml
<ItemGroup>
  <PackageReference Include="DotnetPoi.All" Version="..." />
</ItemGroup>
```

Transitive dependencies (`DotnetPoi.Common`, `DotnetPoi.POIFS`) are resolved automatically by NuGet.

## Why Multiple Packages?

| Package | Contents | Best for |
|---|---|---|
| `DotnetPoi.Ooxml` | XSSF (xlsx/xlsm), XWPF (docx/docm), XSLF (pptx/pptm) | Modern Office 2007+ formats |
| `DotnetPoi.Legacy` | HSSF (xls), HWPF (doc), HSLF (ppt) | Legacy binary formats |
| `DotnetPoi.Formula` | Formula evaluator (`IFormulaEvaluator`, `FormulaEvaluator`, `CellValue`) | Programmatic formula calculation |
| `DotnetPoi.Common` | SS interfaces, shared enums, XML writer | Included transitively |
| `DotnetPoi.POIFS` | OLE2/CFB container, encryption helpers | Included transitively |
| `DotnetPoi.All` | All of the above in one meta-package | Convenience |

**Design principle:** Format packages have zero knowledge of `Formula`. Adding `DotnetPoi.Formula` to your project automatically enables `createFormulaEvaluator()` via lazy assembly discovery at runtime. Without it, the call throws a clear `NotSupportedException`.

## Verify Installation

Create a file `TestInstall.csproj`:

```xml
<Project Sdk="Microsoft.NET.Sdk">
  <PropertyGroup>
    <OutputType>Exe</OutputType>
    <TargetFramework>net8.0</TargetFramework>
  </PropertyGroup>
  <ItemGroup>
    <PackageReference Include="DotnetPoi.Ooxml" Version="..." />
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
