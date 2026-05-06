# Styles and Formatting

dotnet-poi supports fonts, fills, borders, number formats, and alignment. All styles round-trip correctly.

## Creating and Applying a Style

```csharp
using DotnetPoi.XSSF.UserModel;
using DotnetPoi.SS.UserModel;

using var wb = new XSSFWorkbook();

var style = wb.createCellStyle();
var font = wb.createFont();

font.setBold(true);
font.setFontHeightInPoints(14);
font.setFontName("Arial");
font.setColor(IndexedColors.DarkBlue.Index);

style.setFont(font);
style.setAlignment(HorizontalAlignment.Center);

var sheet = wb.createSheet("Styled");
var row = sheet.createRow(0);
var cell = row.createCell(0);
cell.setCellValue("Styled cell");
cell.setCellStyle(style);
```

## Font Properties

| Method | Effect |
|---|---|
| `setBold(bool)` | Bold text |
| `setItalic(bool)` | Italic text |
| `setUnderline(byte)` | Underline style |
| `setStrikeout(bool)` | Strikethrough |
| `setFontName(string)` | Font family (e.g. "Arial", "Times New Roman") |
| `setFontHeightInPoints(int)` | Font size in points |
| `setColor(short)` | Text color (IndexedColors or custom) |

## Cell Fills

```csharp
var style = wb.createCellStyle();
style.setFillForegroundColor(IndexedColors.LightYellow.Index);
style.setFillPattern(FillPatternType.SolidForeground);
```

## Borders

```csharp
style.setBorderBottom(BorderStyle.Thin);
style.setBorderTop(BorderStyle.Medium);
style.setBorderLeft(BorderStyle.Thin);
style.setBorderRight(BorderStyle.Thin);
style.setBottomBorderColor(IndexedColors.DarkBlue.Index);
```

Available border styles: `None`, `Thin`, `Medium`, `Thick`, `Dashed`, `Dotted`, `Double`, and more.

## Number Formats

```csharp
// Built-in formats
style.setDataFormat((short)0);  // General

// Custom format string
style.setDataFormat(wb.createDataFormat().getFormat("#,##0.00"));
style.setDataFormat(wb.createDataFormat().getFormat("yyyy-MM-dd"));
style.setDataFormat(wb.createDataFormat().getFormat("0.0%"));
```

## Alignment

```csharp
style.setAlignment(HorizontalAlignment.Center);       // horizontal
style.setVerticalAlignment(VerticalAlignment.Center); // vertical
style.setWrapText(true);                               // word wrap
style.setIndention(2);                                 // left indent (characters)
style.setRotation(45);                                 // text rotation (degrees)
```

## Full Runnable Example

See `examples/UsageSamples/Program.cs` (`CreateSpreadsheet`):

[examples/UsageSamples](https://github.com/modeverv/dotnet-poi/tree/main/examples/UsageSamples)
