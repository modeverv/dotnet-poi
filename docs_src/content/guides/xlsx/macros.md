# Macros (xlsm)

dotnet-poi preserves VBA macros when round-tripping xlsm files. The `vbaProject.bin` is kept byte-for-byte.

## Preserving Macros on Round-Trip

```csharp
using var stream = File.OpenRead("macro-workbook.xlsm");
using var wb = new XSSFWorkbook(stream);

// Make changes to cells, styles, etc.
var sheet = wb.getSheetAt(0);
sheet.getRow(0).createCell(5).setCellValue("New data");

// Save — macro bytes are preserved automatically
using var outStream = File.Create("macro-workbook-modified.xlsm");
wb.write(outStream);
```

No special handling is needed. The macro parts are automatically preserved as long as you don't remove them.

## Creating Macro-Enabled Files

To create a new xlsm file, write using the `xlsm` extension. The workbook type is determined by the content:

```csharp
using var wb = new XSSFWorkbook();
// ... add content ...
using var file = File.Create("output.xlsm");
wb.write(file);
```

For template-based workflows, start from an existing xlsm that contains the macros you need.

## Limitations

- Creating new VBA macros programmatically is not supported
- Reading VBA project contents is not supported
- Editing VBA macro source code is not supported
- Adding or removing macro references is not supported

## Verified Formats

| Format | Status |
|---|---|
| xlsm (Excel macro-enabled) | ✅ Byte-for-byte preservation verified |
| docm (Word macro-enabled) | ✅ Byte-for-byte preservation verified |
| pptm (PowerPoint macro-enabled) | ✅ Byte-for-byte preservation verified |
