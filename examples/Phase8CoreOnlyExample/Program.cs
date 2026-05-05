using DotnetPoi.SS.UserModel;
using DotnetPoi.XSSF.UserModel;

// ============================================================
// Phase 8: Core-only example — reads and writes xlsx without
//          the DotnetPoi.Formula package.
//
// This demonstrates that DotnetPoi.Core alone is sufficient
// for all spreadsheet read/write/format operations.
//
// Formula evaluation is NOT available (the call to
// createFormulaEvaluator() will throw NotSupportedException).
// ============================================================
#pragma warning disable CS8602

Console.WriteLine("=== Phase 8: DotnetPoi.Core only (no Formula) ===\n");

// ---- Write ----
using (var wb = new XSSFWorkbook())
{
    var sheet = wb.createSheet("Example");

    // --- Row 0: header ---
    var headerFont = wb.createFont();
    headerFont.setFontHeightInPoints(12);
    headerFont.setBold(true);

    var headerStyle = wb.createCellStyle();
    headerStyle.setFont(headerFont);
    headerStyle.setFillForegroundColor((short)IndexedColors.Grey25Percent);
    headerStyle.setFillPattern(FillPatternType.SolidForeground);
    headerStyle.setAlignment(HorizontalAlignment.Center);

    var header = sheet.createRow(0);
    var labels = new[] { "Item", "Price", "Quantity", "Total" };
    for (int i = 0; i < labels.Length; i++)
    {
        var cell = header.createCell(i);
        cell.setCellValue(labels[i]);
        cell.setCellStyle(headerStyle);
    }

    // --- Rows 1-4: data ---
    // Number format for prices
    var priceFmt = wb.createDataFormat();
    var priceStyle = wb.createCellStyle();
    priceStyle.setDataFormat(priceFmt.getFormat("#,##0.00"));

    var qtyStyle = wb.createCellStyle();
    qtyStyle.setAlignment(HorizontalAlignment.Right);

    var totalStyle = wb.createCellStyle();
    totalStyle.setDataFormat(priceFmt.getFormat("#,##0.00"));
    totalStyle.setBorderBottom(BorderStyle.Thin);
    totalStyle.setBorderTop(BorderStyle.Thin);

    var data = new (string Item, double Price, int Qty)[]
    {
        ("Apples",   2.50, 10),
        ("Bananas",  1.20, 20),
        ("Cherries", 4.00,  5),
        ("Dates",    3.75,  8),
    };

    for (int i = 0; i < data.Length; i++)
    {
        var (item, price, qty) = data[i];
        var row = sheet.createRow(i + 1);

        row.createCell(0).setCellValue(item);

        var c1 = row.createCell(1);
        c1.setCellValue(price);
        c1.setCellStyle(priceStyle);

        var c2 = row.createCell(2);
        c2.setCellValue(qty);
        c2.setCellStyle(qtyStyle);

        var c3 = row.createCell(3);
        c3.setCellValue(price * qty);
        c3.setCellStyle(totalStyle);
    }

    // ---- Save to memory stream ----
    using var stream = new MemoryStream();
    wb.write(stream);
    Console.WriteLine($"Written {stream.Length} bytes to memory.\n");

    // ---- Reload from the same stream ----
    stream.Position = 0;
    using var wb2 = new XSSFWorkbook(stream);

    // ---- Read back ----
    var sheet2 = wb2.getSheetAt(0);
    Console.WriteLine($"Sheet index: 0  Rows: {sheet2.getLastRowNum() + 1}\n");

    for (int r = 0; r <= sheet2.getLastRowNum(); r++)
    {
        var row = sheet2.getRow(r);
        if (row == null) continue;

        var parts = new List<string>();
        for (int c = 0; c < 4; c++)
        {
            var cell = row.getCell(c);
            if (cell == null) { parts.Add("(null)"); continue; }

            switch (cell.getCellType())
            {
                case CellType.String:
                    parts.Add(cell.getStringCellValue());
                    break;
                case CellType.Numeric:
                    parts.Add(cell.getNumericCellValue().ToString("F2"));
                    break;
                case CellType.Blank:
                    parts.Add(string.Empty);
                    break;
                default:
                    parts.Add($"[{cell.getCellType()}]");
                    break;
            }
        }
        Console.WriteLine($"  Row {r}: " + string.Join(" | ", parts));
    }

    // ---- Verify styles survive round-trip ----
    var headerCell = sheet2.getRow(0).getCell(0);
    Console.WriteLine($"\n--- Style verification ---");
    Console.WriteLine($"  Header font bold:        {headerCell.getCellStyle().getFont().getBold()}");
    Console.WriteLine($"  Header fill pattern:     {headerCell.getCellStyle().getFillPattern()}");
    Console.WriteLine($"  Header fill color:       {headerCell.getCellStyle().getFillForegroundColor()}");
    Console.WriteLine($"  Header alignment:        {headerCell.getCellStyle().getAlignment()}");

    var totalCell = sheet2.getRow(1).getCell(3);
    Console.WriteLine($"  Total border top:        {totalCell.getCellStyle().getBorderTop()}");
    Console.WriteLine($"  Total border bottom:     {totalCell.getCellStyle().getBorderBottom()}");

    Console.WriteLine($"  Price format string:     \"{sheet2.getRow(1).getCell(1).getCellStyle().getDataFormatString()}\"");

    // ---- Confirm formula evaluator is NOT available ----
    Console.WriteLine($"\n--- Formula evaluator check ---");
    try
    {
        wb2.getCreationHelper().createFormulaEvaluator();
        Console.WriteLine("  ❌ Unexpected: evaluator created (Formula package present?)");
    }
    catch (NotSupportedException ex)
    {
        Console.WriteLine($"  ✅ Expected: {ex.Message}");
    }

    Console.WriteLine("\n=== Done ===");
}
