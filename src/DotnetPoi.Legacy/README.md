# DotnetPoi.Legacy

Legacy Office 97–2003 binary format implementations for **xls** (HSSF), **doc** (HWPF), and **ppt** (HSLF).

## Install

```shell
dotnet add package DotnetPoi.Legacy
```

```xml
<PackageReference Include="DotnetPoi.Legacy" Version="..." />
```

NuGet automatically resolves transitive dependencies (`DotnetPoi.Common`, `DotnetPoi.POIFS`).

## Usage scenarios

| If you need… | Install this |
|---|---|
| xls / doc / ppt read/write only | `DotnetPoi.Legacy` (this package) |
| Legacy binary + formula evaluator | `DotnetPoi.Legacy` + `DotnetPoi.Formula` |
| Legacy binary + OOXML (xlsx/docx/pptx) | `DotnetPoi.Legacy` + `DotnetPoi.Ooxml` |
| Everything (all formats + formula) | `DotnetPoi.All` (meta-package) |

## License

[Apache License 2.0](../../LICENSE) — same as upstream Apache POI.

## Disclaimer

This project is not affiliated with the Apache Software Foundation or the Apache POI project.
Apache POI is a registered trademark of the Apache Software Foundation.
