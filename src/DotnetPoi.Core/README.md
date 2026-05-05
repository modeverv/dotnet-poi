# DotnetPoi.Core

**Faithful .NET port of Apache POI — all format implementations in a single assembly.**

`DotnetPoi.Core` contains every format handler (XSSF/xlsx, HSSF/xls, XWPF/docx, XSLF/pptx, HWPF/doc, HSLF/ppt) plus the POIFS OLE2 container, common SS interfaces, and the `PoiXmlWriter` foundation. It has **zero dependencies** — no formula engine, no external packages.

## When to use

| Scenario | Use |
|---|---|
| Read/write xlsx, xls, docx, pptx, doc, ppt | **DotnetPoi.Core** (always required) |
| + Evaluate spreadsheet formulas | Add **DotnetPoi.Formula** (separate NuGet) |

## What's included

### Formats

| Namespace | Classes | Format |
|---|---|---|
| `DotnetPoi.XSSF.UserModel` | `XSSFWorkbook`, `XSSFSheet`, `XSSFRow`, `XSSFCell`, `XSSFCellStyle`, `XSSFFont`, `XSSFDataFormat`, `XSSFDrawing`, `XSSFPicture`, ... | xlsx (Excel 2007+) |
| `DotnetPoi.HSSF.UserModel` | `HSSFWorkbook`, `HSSFSheet`, `HSSFRow`, `HSSFCell`, `HSSFCellStyle`, `HSSFFont`, ... | xls (Excel 97-2003) |
| `DotnetPoi.XWPF.UserModel` | `XWPFDocument`, `XWPFParagraph`, `XWPFRun`, `XWPFPicture`, `XWPFPictureData` | docx (Word 2007+) |
| `DotnetPoi.XSLF.UserModel` | `XMLSlideShow`, `XSLFSlide`, `XSLFPictureShape`, `XSLFPictureData` | pptx (PowerPoint 2007+) |
| `DotnetPoi.HWPF.UserModel` | `HWPFDocument` (read-only MVP) | doc (Word 97-2003) |
| `DotnetPoi.HSLF.UserModel` | `HSLFSlideShow` (read-only MVP) | ppt (PowerPoint 97-2003) |
| `DotnetPoi.POIFS.FileSystem` | In-repo OLE2 compound document reader/writer (CFB) | OLE2 container |

### Common interfaces

| Interface | Purpose |
|---|---|
| `IWorkbook` / `ISheet` / `IRow` / `ICell` | Unified spreadsheet API across XSSF and HSSF |
| `ICellStyle` / `IFont` / `IDataFormat` | Cell formatting (font, fill, border, alignment, number format) |
| `ICreationHelper` | Factory for data formats, anchors, formula evaluator |

### Infrastructure

- **`PoiXmlWriter`** — XML writer that reproduces Apache POI/XMLBeans output at byte level
- **OOXML Agile Encryption** (AES-128/SHA1) — password-protect `.xlsx` files via `XSSFWorkbook.writeEncrypted()` and decrypt via `EncryptionInfo` / `Encryptor` / `Decryptor` in the `DotnetPoi.POIFS.Crypt` namespace. **docx / pptx encryption is not yet implemented.**

## Quick start

### Project reference

```xml
<ItemGroup>
    <ProjectReference Include="path/to/DotnetPoi.Core/DotnetPoi.Core.csproj" />
</ItemGroup>
```

### NuGet *(when published)*

```xml
<PackageReference Include="DotnetPoi.Core" Version="1.0.0" />
```

### xlsx write example

```csharp
using DotnetPoi.XSSF.UserModel;

using var wb = new XSSFWorkbook();
var sheet = wb.createSheet("Sheet1");
var row = sheet.createRow(0);
row.createCell(0).setCellValue("Hello");
row.createCell(1).setCellValue(42);

using var fs = new FileStream("output.xlsx", FileMode.Create);
wb.write(fs);
```

### xlsx read example

```csharp
using var fs = new FileStream("input.xlsx", FileMode.Open);
using var wb = new XSSFWorkbook(fs);

var sheet = wb.getSheetAt(0);
var cell = sheet.getRow(0).getCell(0);
Console.WriteLine(cell.getStringCellValue()); // "Hello"
```

### docx write example

```csharp
using DotnetPoi.XWPF.UserModel;

using var doc = new XWPFDocument();
var para = doc.createParagraph();
var run = para.createRun();
run.setText("dotnet-poi can write docx");
run.setBold(true);

using var fs = new FileStream("output.docx", FileMode.Create);
doc.write(fs);
```

### pptx write example

```csharp
using DotnetPoi.XSLF.UserModel;

using var prs = new XMLSlideShow();
var slide = prs.createSlide();
var picIdx = prs.addPicture(File.ReadAllBytes("photo.jpeg"), XSLFPictureData.PICTURE_TYPE_JPEG);
var shape = prs.createPicture(slide, picIdx);
shape.setAnchor(0, 0, XMLSlideShow.DefaultSlideCx, XMLSlideShow.DefaultSlideCy);
shape.setRotation(45.0);

using var fs = new FileStream("output.pptx", FileMode.Create);
prs.write(fs);
```

### Encrypted xlsx write example (xlsx only — docx/pptx encryption not yet implemented)

```csharp
using DotnetPoi.XSSF.UserModel;

using var wb = new XSSFWorkbook();
var sheet = wb.createSheet("Sheet1");
sheet.createRow(0).createCell(0).setCellValue("Encrypted content");

using var fs = new FileStream("protected.xlsx", FileMode.Create);
wb.writeEncrypted(fs, "password123");
// Opens in Excel with the password "password123"
```

### Decrypt and read example

```csharp
using DotnetPoi.POIFS.Crypt;
using DotnetPoi.XSSF.UserModel;

using var fs = new FileStream("protected.xlsx", FileMode.Open);
var info = new EncryptionInfo(fs);

var decryptor = info.Decryptor;
if (decryptor.verifyPassword("password123"))
{
    var decryptedData = decryptor.getData();
    using var ms = new MemoryStream(decryptedData);
    using var wb = new XSSFWorkbook(ms);
    Console.WriteLine(wb.getSheetAt(0).getRow(0).getCell(0).getStringCellValue());
}
```

Encryption uses OOXML Agile Encryption (AES-128/SHA1, 100,000 spin count), compatible with Apache POI Java and Microsoft Excel.

## Architecture note

`DotnetPoi.Core` contains **everything except formula evaluation**. Formulas are preserved as text and cached values are read/written, but the evaluator engine lives in the separate `DotnetPoi.Formula` package.

When `DotnetPoi.Formula` is referenced in your project, `createFormulaEvaluator()` automatically discovers it at runtime via lazy assembly reflection. Without it, the call throws `NotSupportedException` with a clear message.

See `DotnetPoi.Formula/README.md` for the supported formula functions list.

## License

[Apache License 2.0](../../LICENSE) — same as upstream Apache POI.
