# DotnetPoi.All

Meta-package that bundles the complete DotnetPoi surface under a single dependency:

- **DotnetPoi.Ooxml** — xlsx, docx, pptx (OOXML formats)
- **DotnetPoi.Legacy** — xls, doc, ppt (legacy binary formats)
- **DotnetPoi.Formula** — formula evaluator (limited subset)
- **DotnetPoi.Common** — shared interfaces and utilities
- **DotnetPoi.POIFS** — OLE2 / CFB container foundation

## Install

```shell
dotnet add package DotnetPoi.All
```

```xml
<PackageReference Include="DotnetPoi.All" Version="..." />
```

## When to use

| Use this if you… | Otherwise consider… |
|---|---|
| Want everything with one dependency | `DotnetPoi.Ooxml` for OOXML-only projects |
| Don't want to think about granular package selection | `DotnetPoi.Legacy` for legacy-only projects |
| Need all formats + formula evaluation | `DotnetPoi.Ooxml` + `DotnetPoi.Formula` for minimal footprint |

## Current scope

- OOXML (`xlsx`, `docx`, `pptx`) is the most practical surface. `docx` tracked-change XML is preserved during round-trip, but revision accept/reject/create/edit APIs are not modeled.
- Legacy binary formats are included for convenience, but remain partial: `xls` and `doc` cover practical read/write/light-edit slices, while `ppt` is preservation/text-extraction oriented.
- Formula evaluation is intentionally limited. Formula text and cached-value preservation live in the format packages; full Excel-compatible calculation is not a project goal.

## License

[Apache License 2.0](../../LICENSE) — same as upstream Apache POI.

## Disclaimer

This project is not affiliated with the Apache Software Foundation or the Apache POI project.
Apache POI is a registered trademark of the Apache Software Foundation.
