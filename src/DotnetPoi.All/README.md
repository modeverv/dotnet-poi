# DotnetPoi.All

Meta-package that bundles the DotnetPoi packages under a single dependency:

- **DotnetPoi.Ooxml 1.0** — stable xlsx, docx, pptx OOXML workflows
- **DotnetPoi.Legacy 0.x** — partial xls, doc, ppt legacy binary support
- **DotnetPoi.Formula 0.x** — formula evaluator for a limited supported subset
- **DotnetPoi.Common** — shared interfaces and utilities
- **DotnetPoi.POIFS** — OLE2 / CFB container foundation

`DotnetPoi.All` 1.0 is a convenience package. Its 1.0 stability promise comes from `DotnetPoi.Ooxml`; Legacy and Formula are included for convenience but remain partial and documented as such.

## Install

```shell
dotnet add package DotnetPoi.All --version 1.0.0
```

```xml
<PackageReference Include="DotnetPoi.All" Version="1.0.0" />
```

## When to use

| Use this if you… | Otherwise consider… |
|---|---|
| Want everything with one dependency | `DotnetPoi.Ooxml` for OOXML-only projects |
| Don't want to think about granular package selection | `DotnetPoi.Legacy` for legacy-only projects |
| Need all formats + formula evaluation | `DotnetPoi.Ooxml` + `DotnetPoi.Formula` for minimal footprint |

## Support note

Use `DotnetPoi.Ooxml` directly when you only need modern Office files and want the smallest stable dependency. Use `DotnetPoi.All` when the convenience of one package is more important than separating stable OOXML from in-development Legacy and Formula surfaces.

## License

[Apache License 2.0](../../LICENSE) — same as upstream Apache POI.

## Disclaimer

This project is not affiliated with the Apache Software Foundation or the Apache POI project.
Apache POI is a registered trademark of the Apache Software Foundation.
