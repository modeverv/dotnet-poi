# DotnetPoi.Ooxml

OOXML implementations for **xlsx**, **docx**, **pptx**, and macro-enabled Office 2007+ formats.

## Install

```shell
dotnet add package DotnetPoi.Ooxml
```

```xml
<PackageReference Include="DotnetPoi.Ooxml" Version="..." />
```

NuGet automatically resolves transitive dependencies (`DotnetPoi.Common`, `DotnetPoi.POIFS`).

## Usage scenarios

| If you need… | Install this |
|---|---|
| xlsx / docx / pptx read/write only | `DotnetPoi.Ooxml` (this package) |
| OOXML + formula evaluator | `DotnetPoi.Ooxml` + `DotnetPoi.Formula` |
| OOXML + legacy binary (xls/doc/ppt) | `DotnetPoi.Ooxml` + `DotnetPoi.Legacy` |
| Everything (all formats + formula) | `DotnetPoi.All` (meta-package) |

## License

[Apache License 2.0](../../LICENSE) — same as upstream Apache POI.

## Disclaimer

This project is not affiliated with the Apache Software Foundation or the Apache POI project.
Apache POI is a registered trademark of the Apache Software Foundation.
