using System.Buffers.Binary;
using System.Text;
using System.Linq;
using System.Collections.Generic;
using DotnetPoi.HSSF.UserModel;
using DotnetPoi.SS.UserModel;

namespace DotnetPoi.HSSF.Record;

internal static class Biff8Workbook
{
    private const ushort Bof = 0x0809;
    private const ushort Eof = 0x000A;
    private const ushort BoundSheet8 = 0x0085;
    private const ushort Sst = 0x00FC;
    private const ushort Label = 0x0204;
    private const ushort LabelSst = 0x00FD;
    private const ushort Number = 0x0203;
    private const ushort Rk = 0x027E;
    private const ushort BoolErr = 0x0205;
    private const ushort Blank = 0x0201;
    private const ushort Dimensions = 0x0200;
    private const ushort Window2 = 0x023E;
    private const ushort Selection = 0x001D;
    private const ushort CodePage = 0x0042;
    private const ushort Window1 = 0x003D;
    private const ushort FontRecord = 0x0031;
    private const ushort FormatRecord = 0x041E;
    private const ushort XfRecord = 0x00E0;
    private const ushort RowRecord = 0x0208;
    private const ushort ColInfo = 0x007D;
    private const ushort MergeCellsRecord = 0x00E5;
    private const ushort PaneRecord = 0x0041;
    private const ushort StyleRecord = 0x0293;

    public static void ReadWorkbook(byte[] workbookStream, HSSFWorkbook workbook)
    {
        var records = ReadRecords(workbookStream);
        var sheets = new List<BoundSheet>();
        var sharedStrings = new List<string>();
        var fontCount = 0;
        var xfCount = 0;

        foreach (var record in records)
        {
            switch (record.Sid)
            {
                case FontRecord:
                    // Font index 4 is reserved in BIFF8 — skip it by adding a placeholder
                    if (fontCount == 4) { workbook.AddFontFromBiff(new HSSFFont(4)); fontCount++; }
                    workbook.AddFontFromBiff(ReadFont(record.Data, fontCount));
                    fontCount++;
                    break;
                case FormatRecord:
                    ReadFormatRecord(workbook, record.Data);
                    break;
                case XfRecord:
                    // Store all XF records in HSSFWorkbook with 1:1 index mapping
                    workbook.AddStyleFromBiff(ReadXf(workbook, record.Data, xfCount));
                    xfCount++;
                    break;
                case BoundSheet8:
                    sheets.Add(ReadBoundSheet(record.Data));
                    break;
                case Sst:
                    sharedStrings = ReadSst(record.Data);
                    break;
            }
        }

        if (sheets.Count == 0)
        {
            workbook.createSheet();
            return;
        }

        foreach (var boundSheet in sheets)
        {
            var sheet = workbook.createSheet(boundSheet.Name);
            ReadSheet(records, boundSheet.Offset, sheet, sharedStrings);
        }
    }

    private static void ReadFormatRecord(HSSFWorkbook workbook, ReadOnlySpan<byte> data)
    {
        if (data.Length < 5) return;
        var index = BinaryPrimitives.ReadInt16LittleEndian(data);
        var len = BinaryPrimitives.ReadUInt16LittleEndian(data.Slice(2));
        var unicode = data[4] != 0;
        var byteCount = len * (unicode ? 2 : 1);
        if (data.Length < 5 + byteCount) return;
        var format = unicode
            ? Encoding.Unicode.GetString(data.Slice(5, byteCount).ToArray())
            : Encoding.GetEncoding("ISO-8859-1").GetString(data.Slice(5, byteCount).ToArray());
        workbook.GetOrCreateDataFormat().AddBiffFormat(index, format);
    }

    private static readonly (short index, string format)[] BuiltinFormats = new[]
    {
        ((short)5, "\"$\"#,##0_);(\"$\"#,##0)"),
        ((short)6, "\"$\"#,##0_);[Red](\"$\"#,##0)"),
        ((short)7, "\"$\"#,##0.00_);(\"$\"#,##0.00)"),
        ((short)8, "\"$\"#,##0.00_);[Red](\"$\"#,##0.00)"),
        ((short)42, "_(\"$\"* #,##0_);_(\"$\"* (#,##0);_(\"$\"* \"-\"_);_(@_)"),
        ((short)41, "_(* #,##0_);_(* (#,##0);_(* \"-\"_);_(@_)"),
        ((short)44, "_(\"$\"* #,##0.00_);_(\"$\"* (#,##0.00);_(\"$\"* \"-\"??_);_(@_)"),
        ((short)43, "_(* #,##0.00_);_(* (#,##0.00);_(* \"-\"??_);_(@_)")
    };

    private static void WriteFormatRecords(Stream stream, HSSFWorkbook workbook)
    {
        // Write 8 built-in FormatRecord entries (matching POI)
        foreach (var (idx, format) in BuiltinFormats)
        {
            WriteFormatRecord(stream, idx, format);
        }

        // Write user-defined FormatRecord entries (index >= 164)
        foreach (var kv in workbook.GetOrCreateDataFormat().GetUserDefinedFormats())
        {
            WriteFormatRecord(stream, kv.Key, kv.Value);
        }
    }

    private static void WriteFormatRecord(Stream stream, short index, string format)
    {
        WriteRecord(stream, FormatRecord, payload =>
        {
            Span<byte> header = stackalloc byte[4];
            BinaryPrimitives.WriteInt16LittleEndian(header, index);
            BinaryPrimitives.WriteUInt16LittleEndian(header.Slice(2), (ushort)Math.Min(format.Length, ushort.MaxValue));
            payload.Write(header.ToArray(), 0, header.Length);
            var compressed = format.All(ch => ch <= 0x00FF);
            payload.WriteByte(compressed ? (byte)0 : (byte)1);
            var bytes = compressed
                ? Encoding.GetEncoding("ISO-8859-1").GetBytes(format)
                : Encoding.Unicode.GetBytes(format);
            payload.Write(bytes, 0, bytes.Length);
        });
    }

    private static HSSFFont ReadFont(ReadOnlySpan<byte> data, int index)
    {
        var font = new HSSFFont(index);
        if (data.Length < 14) return font;

        var height = BinaryPrimitives.ReadInt16LittleEndian(data);
        var attributes = BinaryPrimitives.ReadInt16LittleEndian(data.Slice(2));
        var color = BinaryPrimitives.ReadInt16LittleEndian(data.Slice(4));
        var boldWeight = BinaryPrimitives.ReadInt16LittleEndian(data.Slice(6));
        var underline = data.Length > 10 ? data[10] : (byte)0;

        font.setFontHeight(height);
        font.setColor(color);
        font.setBold(boldWeight >= 700);
        font.setItalic((attributes & 0x02) != 0);
        font.setStrikeout((attributes & 0x08) != 0);
        font.setUnderline(underline);

        var nameLen = data.Length > 14 ? data[14] : (byte)0;
        var unicodeFlags = data.Length > 15 ? data[15] : (byte)0;

        if (nameLen > 0)
        {
            var byteCount = nameLen * (unicodeFlags == 0 ? 1 : 2);
            if (data.Length >= 16 + byteCount)
            {
                font.setFontName(unicodeFlags == 0
                    ? Encoding.GetEncoding("ISO-8859-1").GetString(data.Slice(16, byteCount).ToArray())
                    : Encoding.Unicode.GetString(data.Slice(16, byteCount).ToArray()));
            }
        }

        return font;
    }

    private static HSSFCellStyle ReadXf(HSSFWorkbook workbook, ReadOnlySpan<byte> data, int styleIndex)
    {
        var style = new HSSFCellStyle(workbook, styleIndex);
        if (data.Length < 20) return style;

        var fontIndex = BinaryPrimitives.ReadUInt16LittleEndian(data);
        var formatIndex = BinaryPrimitives.ReadInt16LittleEndian(data.Slice(2));
        var cellOptions = BinaryPrimitives.ReadUInt16LittleEndian(data.Slice(4));
        var alignmentOptions = BinaryPrimitives.ReadUInt16LittleEndian(data.Slice(6));
        var borderOptions = BinaryPrimitives.ReadUInt16LittleEndian(data.Slice(10));
        var fillPaletteOptions = BinaryPrimitives.ReadUInt16LittleEndian(data.Slice(18));

        if ((cellOptions & 0x0004) == 0) // Cell XF
        {
            style.setParentIndex((short)(cellOptions >> 4));
        }

        style.setDataFormat(formatIndex);

        var hAlignVal = alignmentOptions & 0x07;
        if (Enum.IsDefined(typeof(HorizontalAlignment), hAlignVal)) style.setAlignment((HorizontalAlignment)hAlignVal);
        style.setWrapText((alignmentOptions & 0x08) != 0);
        var vAlignVal = (short)((alignmentOptions >> 4) & 0x07);
        if (Enum.IsDefined(typeof(VerticalAlignment), vAlignVal)) style.setVerticalAlignment((VerticalAlignment)vAlignVal);

        var borderLeft = (short)(borderOptions & 0x0F);
        var borderRight = (short)((borderOptions >> 4) & 0x0F);
        var borderTop = (short)((borderOptions >> 8) & 0x0F);
        var borderBottom = (short)((borderOptions >> 12) & 0x0F);
        if (Enum.IsDefined(typeof(BorderStyle), borderLeft)) style.setBorderLeft((BorderStyle)borderLeft);
        if (Enum.IsDefined(typeof(BorderStyle), borderRight)) style.setBorderRight((BorderStyle)borderRight);
        if (Enum.IsDefined(typeof(BorderStyle), borderTop)) style.setBorderTop((BorderStyle)borderTop);
        if (Enum.IsDefined(typeof(BorderStyle), borderBottom)) style.setBorderBottom((BorderStyle)borderBottom);

        var fillPattern = (fillPaletteOptions >> 10) & 0x3F;
        if (Enum.IsDefined(typeof(FillPatternType), fillPattern)) style.setFillPattern((FillPatternType)fillPattern);
        style.setFillForegroundColor((short)((fillPaletteOptions >> 7) & 0x7F));

        if (fontIndex < workbook.getNumberOfFonts())
            style.setFont(workbook.getFontAt((int)fontIndex));

        return style;
    }

    public static byte[] WriteWorkbook(IReadOnlyList<HSSFSheet> sheets, byte[]? templateWorkbookStream = null)
    {
        return templateWorkbookStream is null
            ? WriteNewWorkbook(sheets)
            : WriteWorkbookPreservingRecords(sheets, templateWorkbookStream);
    }

    private static byte[] WriteNewWorkbook(IReadOnlyList<HSSFSheet> sheets)
    {
        var workbook = sheets.Count > 0 ? sheets[0].getWorkbook() : new HSSFWorkbook();

        using var stream = new MemoryStream();
        WriteBof(stream, 0x0005);
        WriteRecord(stream, CodePage, payload =>
        {
            Span<byte> buffer = stackalloc byte[2];
            BinaryPrimitives.WriteUInt16LittleEndian(buffer, 1200);
            payload.Write(buffer.ToArray(), 0, buffer.Length);
        });
        WriteRecord(stream, Window1, payload =>
        {
            Span<byte> buffer = stackalloc byte[18];
            buffer.Clear();
            BinaryPrimitives.WriteUInt16LittleEndian(((Span<byte>)buffer).Slice(8), 0x38);
            payload.Write(buffer.ToArray(), 0, buffer.Length);
        });

        // Write Font records: 4 defaults + user-defined fonts
        WriteFontRecords(stream, workbook);

        // Write user-defined FormatRecord entries (index >= 164)
        WriteFormatRecords(stream, workbook);

        // Write 21 built-in XF + user cell XF records
        WriteXfRecords(stream, workbook);

        // Write 6 built-in StyleRecord entries (matching POI)
        WriteStyleRecords(stream);

        var boundSheetOffsetPositions = new List<long>();
        foreach (var sheet in sheets.Count == 0 ? new[] { new HSSFSheet(new HSSFWorkbook(), "Sheet1") } : sheets)
        {
            boundSheetOffsetPositions.Add(stream.Position + 4);
            WriteBoundSheet(stream, sheet.SheetName, 0);
        }

        var strings = BuildSharedStrings(sheets);
        WriteSst(stream, strings);
        WriteRecord(stream, Eof, _ => { });

        var sheetIndex = 0;
        foreach (var sheet in sheets.Count == 0 ? Array.Empty<HSSFSheet>() : sheets)
        {
            var sheetOffset = checked((uint)stream.Position);
            PatchUInt32(stream, boundSheetOffsetPositions[sheetIndex++], sheetOffset);
            WriteSheet(stream, sheet, strings);
        }

        if (sheets.Count == 0)
        {
            PatchUInt32(stream, boundSheetOffsetPositions[0], checked((uint)stream.Position));
            WriteSheet(stream, new HSSFSheet(new HSSFWorkbook(), "Sheet1"), strings);
        }

        return stream.ToArray();
    }

    private static void WriteFontRecords(Stream stream, HSSFWorkbook workbook)
    {
        // Always write 4 default fonts (BIFF8 requirement); user fonts start at index 5 (index 4 reserved)
        var numFonts = workbook.getNumberOfFonts();
        for (var i = 0; i < 4; i++)
        {
            var font = i < numFonts ? workbook.getFontAt(i) : null;
            WriteFontRecord(stream, font);
        }
        // User fonts start at index 4 in our list (placeholder) → BIFF font index 5 (skip index 4)
        for (var i = 4; i < numFonts; i++)
        {
            if (i == 4) continue; // Skip index 4 placeholder
            WriteFontRecord(stream, workbook.getFontAt(i));
        }
    }

    private static void WriteFontRecord(Stream stream, HSSFFont? font)
    {
        WriteRecord(stream, FontRecord, payload =>
        {
            var height = font?.getFontHeight() ?? (short)200;
            var attributes = font?.Attributes ?? (short)0;
            var color = font?.AutoColor ?? (short)0x7FFF;
            var boldWeight = font?.BoldWeight ?? (short)400;
            var underline = font?.getUnderline() ?? (byte)0;
            var name = font?.getFontName() ?? "Arial";
            if (string.IsNullOrEmpty(name)) name = "Arial";

            Span<byte> fixed14 = stackalloc byte[14];
            BinaryPrimitives.WriteInt16LittleEndian(fixed14, height);
            BinaryPrimitives.WriteInt16LittleEndian(fixed14.Slice(2), attributes);
            BinaryPrimitives.WriteInt16LittleEndian(fixed14.Slice(4), color);
            BinaryPrimitives.WriteInt16LittleEndian(fixed14.Slice(6), boldWeight);
            fixed14.Slice(8).Clear(); // super_sub=0, underline will be set below
            fixed14[10] = underline;
            payload.Write(fixed14.ToArray(), 0, fixed14.Length);

            var compressed = name.All(ch => ch <= 0x00FF);
            payload.WriteByte((byte)Math.Min(name.Length, 255));
            payload.WriteByte(compressed ? (byte)0 : (byte)1);
            var nameBytes = compressed
                ? Encoding.GetEncoding("ISO-8859-1").GetBytes(name)
                : Encoding.Unicode.GetBytes(name);
            payload.Write(nameBytes, 0, nameBytes.Length);
        });
    }

    // 21 built-in XF records (hard-coded to match POI defaults)
    private static readonly byte[][] BuiltinStyleXfData = GenerateBuiltinStyleXf();

    private static byte[][] GenerateBuiltinStyleXf()
    {
        var result = new byte[21][];
        for (var i = 0; i < 21; i++)
        {
            int fontIndex, formatIndex, cellOptions, indentionOptions;
            switch (i)
            {
                case 0: fontIndex = 0; formatIndex = 0; cellOptions = 0xFFF5; indentionOptions = 0; break;
                case 1:
                case 2: fontIndex = 1; formatIndex = 0; cellOptions = 0xFFF5; indentionOptions = 0xF400; break;
                case 3:
                case 4: fontIndex = 2; formatIndex = 0; cellOptions = 0xFFF5; indentionOptions = 0xF400; break;
                case 5:
                case 6:
                case 7:
                case 8:
                case 9:
                case 10:
                case 11:
                case 12:
                case 13:
                case 14: fontIndex = 0; formatIndex = 0; cellOptions = 0xFFF5; indentionOptions = 0xF400; break;
                case 15: fontIndex = 0; formatIndex = 0; cellOptions = 0x0001; indentionOptions = 0; break;
                case 16: fontIndex = 1; formatIndex = 0x2b; cellOptions = 0xFFF5; indentionOptions = 0xF800; break;
                case 17: fontIndex = 1; formatIndex = 0x29; cellOptions = 0xFFF5; indentionOptions = 0xF800; break;
                case 18: fontIndex = 1; formatIndex = 0x2c; cellOptions = 0xFFF5; indentionOptions = 0xF800; break;
                case 19: fontIndex = 1; formatIndex = 0x2a; cellOptions = 0xFFF5; indentionOptions = 0xF800; break;
                case 20: fontIndex = 1; formatIndex = 0x09; cellOptions = 0xFFF5; indentionOptions = 0xF800; break;
                default: throw new ArgumentException("Unrecognized format id: " + i);
            }

            var data = new byte[20];
            var span = (Span<byte>)data;
            BinaryPrimitives.WriteUInt16LittleEndian(span, (ushort)fontIndex);
            BinaryPrimitives.WriteUInt16LittleEndian(span.Slice(2), (ushort)formatIndex);
            BinaryPrimitives.WriteUInt16LittleEndian(span.Slice(4), (ushort)cellOptions);
            BinaryPrimitives.WriteUInt16LittleEndian(span.Slice(6), 0x20); // alignment
            BinaryPrimitives.WriteUInt16LittleEndian(span.Slice(8), (ushort)indentionOptions);
            data[18] = 0xC0; data[19] = 0x20; // fill = 0x20C0
            result[i] = data;
        }
        return result;
    }

    private static void WriteXfRecords(Stream stream, HSSFWorkbook workbook)
    {
        // Write 21 built-in XF records
        foreach (var data in BuiltinStyleXfData)
        {
            WriteRecord(stream, XfRecord, payload => payload.Write(data, 0, data.Length));
        }

        // Write user cell XF records starting from index 21
        var numStyles = workbook.getNumberOfCellStyles();
        for (var i = 21; i < numStyles; i++)
        {
            var style = workbook.getCellStyleAt(i);
            WriteXfRecord(stream, style);
        }
    }

    private static void WriteStyleRecords(Stream stream)
    {
        // 6 built-in styles: Currency, Currency [0], Comma, Comma [0], Normal, Percent
        // Hard-coded bytes for built-in styles (XFIndex | 0x8000, BuiltinStyle, OutlineStyleLevel)
        byte[][] styles = new byte[][]
        {
            new byte[] { 0x10, 0x80, 0x03, 0xFF }, // Currency
            new byte[] { 0x11, 0x80, 0x06, 0xFF }, // Currency [0]
            new byte[] { 0x12, 0x80, 0x04, 0xFF }, // Comma
            new byte[] { 0x13, 0x80, 0x07, 0xFF }, // Comma [0]
            new byte[] { 0x00, 0x80, 0x00, 0xFF }, // Normal
            new byte[] { 0x14, 0x80, 0x05, 0xFF }  // Percent
        };

        foreach (var data in styles)
        {
            WriteRecord(stream, StyleRecord, payload => payload.Write(data, 0, data.Length));
        }
    }

    private static void WriteXfRecord(Stream stream, HSSFCellStyle style)
    {
        WriteRecord(stream, XfRecord, payload =>
        {
            Span<byte> data = stackalloc byte[20];
            data.Clear();

            BinaryPrimitives.WriteUInt16LittleEndian(data, (ushort)style.FontBiffIndex);
            BinaryPrimitives.WriteInt16LittleEndian(data.Slice(2), style.getDataFormat());
            data[4] = 0x01; // cell_options = 0x0001 (cell XF type)

            var alignOpts = (ushort)((int)style.getAlignment() & 0x07);
            if (style.getWrapText()) alignOpts |= 0x08;
            alignOpts |= (ushort)(((int)style.getVerticalAlignment() & 0x07) << 4);
            BinaryPrimitives.WriteUInt16LittleEndian(data.Slice(6), alignOpts);

            // indention_options = 0
            // border_options
            var borderOpts = (ushort)((int)style.getBorderLeft() & 0x0F);
            borderOpts |= (ushort)(((int)style.getBorderRight() & 0x0F) << 4);
            borderOpts |= (ushort)(((int)style.getBorderTop() & 0x0F) << 8);
            borderOpts |= (ushort)(((int)style.getBorderBottom() & 0x0F) << 12);
            BinaryPrimitives.WriteUInt16LittleEndian(data.Slice(10), borderOpts);

            // adtl_palette_options = 0 (4 bytes, bytes 14-17)
            // fill_palette_options: bits 10-15=fillPattern, bits 7-13=fgColor, bits 0-6=bgColor
            var fillOpts = (ushort)(((int)style.getFillPattern() << 10) & 0xFC00);
            fillOpts |= (ushort)((style.getFillForegroundColor() & 0x7F) << 7);
            fillOpts |= 0x0040; // background color = 64 (AUTOMATIC)
            BinaryPrimitives.WriteUInt16LittleEndian(data.Slice(18), fillOpts);

            payload.Write(data.ToArray(), 0, data.Length);
        });
    }

    private static byte[] WriteWorkbookPreservingRecords(IReadOnlyList<HSSFSheet> sheets, byte[] templateWorkbookStream)
    {
        var workbook = sheets.Count > 0 ? sheets[0].getWorkbook() : new HSSFWorkbook();
        var records = ReadRecords(templateWorkbookStream);
        var boundSheets = records
            .Where(record => record.Sid == BoundSheet8)
            .Select(record => ReadBoundSheet(record.Data))
            .ToArray();
        if (boundSheets.Length != sheets.Count || sheets.Count == 0)
        {
            return WriteNewWorkbook(sheets);
        }

        var firstSheetOffset = boundSheets.Min(sheet => sheet.Offset);
        var globalRecords = records.TakeWhile(record => record.Offset < firstSheetOffset).ToArray();
        var existingFontCount = globalRecords.Count(r => r.Sid == FontRecord);
        var existingXfCount = globalRecords.Count(r => r.Sid == XfRecord);
        var existingFormatCount = globalRecords.Count(r => r.Sid == FormatRecord);
        var existingFormatIndices = new HashSet<short>(globalRecords
            .Where(r => r.Sid == FormatRecord && r.Data.Length >= 2)
            .Select(r => BinaryPrimitives.ReadInt16LittleEndian(r.Data)));

        var totalFonts = workbook.getNumberOfFonts();
        var totalStyles = workbook.getNumberOfCellStyles();

        IEnumerable<string> existingStrings = records.FirstOrDefault(record => record.Sid == Sst) is { } sstRecord
            ? ReadSst(sstRecord.Data)
            : Array.Empty<string>();
        var strings = BuildSharedStrings(sheets, existingStrings);

        using var stream = new MemoryStream();
        var boundSheetOffsetPositions = new List<long>();
        var wroteBoundSheets = false;
        var wroteSst = false;
        var fontCount = 0;
        var xfCount = 0;
        var formatCount = 0;

        foreach (var record in globalRecords)
        {
            switch (record.Sid)
            {
                case FontRecord:
                    WriteRawRecord(stream, record);
                    fontCount++;
                    if (fontCount == existingFontCount)
                    {
                        var startIdx = existingFontCount + (existingFontCount >= 4 ? 1 : 0);
                        for (var i = startIdx; i < totalFonts; i++)
                        {
                            WriteFontRecord(stream, workbook.getFontAt(i));
                        }
                    }
                    break;
                case FormatRecord:
                    WriteRawRecord(stream, record);
                    formatCount++;
                    if (formatCount == existingFormatCount)
                    {
                        foreach (var kv in workbook.GetOrCreateDataFormat().GetUserDefinedFormats())
                        {
                            if (!existingFormatIndices.Contains(kv.Key))
                            {
                                WriteFormatRecord(stream, kv.Key, kv.Value);
                            }
                        }
                    }
                    break;
                case XfRecord:
                    WriteRawRecord(stream, record);
                    xfCount++;
                    if (xfCount == existingXfCount)
                    {
                        // Append new user-defined styles
                        for (var i = xfCount; i < totalStyles; i++)
                        {
                            WriteXfRecord(stream, workbook.getCellStyleAt(i));
                        }
                    }
                    break;
                case BoundSheet8 when !wroteBoundSheets:
                    foreach (var sheet in sheets)
                    {
                        boundSheetOffsetPositions.Add(stream.Position + 4);
                        WriteBoundSheet(stream, sheet.SheetName, 0);
                    }

                    wroteBoundSheets = true;
                    break;
                case BoundSheet8:
                    break;
                case Sst when !wroteSst:
                    WriteSst(stream, strings);
                    wroteSst = true;
                    break;
                case Sst:
                    break;
                case Eof:
                    if (!wroteBoundSheets)
                    {
                        foreach (var sheet in sheets)
                        {
                            boundSheetOffsetPositions.Add(stream.Position + 4);
                            WriteBoundSheet(stream, sheet.SheetName, 0);
                        }

                        wroteBoundSheets = true;
                    }

                    if (fontCount == 0 && totalFonts > 0)
                    {
                         // No Font records found in template? (should not happen in valid BIFF8)
                         var startIdx = 0;
                         for (var i = startIdx; i < totalFonts; i++)
                         {
                             if (i == 4) continue;
                             WriteFontRecord(stream, workbook.getFontAt(i));
                         }
                    }

                    if (formatCount == 0)
                    {
                        foreach (var kv in workbook.GetOrCreateDataFormat().GetUserDefinedFormats())
                        {
                            WriteFormatRecord(stream, kv.Key, kv.Value);
                        }
                    }

                    if (xfCount == 0 && totalStyles > 0)
                    {
                         // No XF records found in template? (should not happen)
                         for (var i = 0; i < totalStyles; i++) WriteXfRecord(stream, workbook.getCellStyleAt(i));
                    }

                    if (!wroteSst)
                    {
                        WriteSst(stream, strings);
                    }

                    WriteRawRecord(stream, record);
                    break;
                default:
                    WriteRawRecord(stream, record);
                    break;
            }
        }

        if (boundSheetOffsetPositions.Count != sheets.Count)
        {
            return WriteNewWorkbook(sheets);
        }

        for (var i = 0; i < sheets.Count; i++)
        {
            var sheetOffset = checked((uint)stream.Position);
            PatchUInt32(stream, boundSheetOffsetPositions[i], sheetOffset);
            WriteSheetPreservingRecords(stream, sheets[i], strings, GetSheetRecords(records, boundSheets[i].Offset));
        }

        return stream.ToArray();
    }

    private static void ReadSheet(IReadOnlyList<Record> records, uint offset, HSSFSheet sheet, IReadOnlyList<string> sharedStrings)
    {
        var workbook = sheet.getWorkbook();
        var start = 0;
        while (start < records.Count && records[start].Offset < offset)
        {
            start++;
        }

        for (var i = start; i < records.Count; i++)
        {
            var record = records[i];
            if (record.Offset > offset && record.Sid == Bof && i != start)
            {
                break;
            }

            if (record.Offset >= offset && record.Sid == Eof)
            {
                break;
            }

            switch (record.Sid)
            {
                case RowRecord:
                    ReadRowRecord(sheet, record.Data);
                    break;
                case ColInfo:
                    ReadColInfo(sheet, record.Data);
                    break;
                case MergeCellsRecord:
                    ReadMergeCells(sheet, record.Data);
                    break;
                case PaneRecord:
                    ReadPaneRecord(sheet, record.Data);
                    break;
                case LabelSst:
                    ReadLabelSst(workbook, sheet, record.Data, sharedStrings);
                    break;
                case Label:
                    ReadLabel(workbook, sheet, record.Data);
                    break;
                case Number:
                    ReadNumber(workbook, sheet, record.Data);
                    break;
                case Rk:
                    ReadRk(workbook, sheet, record.Data);
                    break;
                case BoolErr:
                    ReadBoolErr(workbook, sheet, record.Data);
                    break;
                case Blank:
                    ReadBlank(workbook, sheet, record.Data);
                    break;
            }
        }
    }

    private static void SetCellStyleFromXfIndex(HSSFWorkbook workbook, HSSFCell cell, int xfIndex)
    {
        // 15 is the default cell XF (built-in).
        if (xfIndex == 15)
        {
            cell.setCellStyle(null); // Use default style
            return;
        }

        if (xfIndex >= 0 && xfIndex < workbook.getNumberOfCellStyles())
        {
            cell.setCellStyle(workbook.getCellStyleAt(xfIndex));
        }
    }

    private static void ReadRowRecord(HSSFSheet sheet, ReadOnlySpan<byte> data)
    {
        if (data.Length < 16) return;
        var rowIndex = BinaryPrimitives.ReadUInt16LittleEndian(data);
        var heightTwips = BinaryPrimitives.ReadUInt16LittleEndian(data.Slice(6));
        var optionFlags = BinaryPrimitives.ReadUInt16LittleEndian(data.Slice(12));
        var hidden = (optionFlags & 0x0020) != 0; // zeroHeight bit

        var row = sheet.getRow(rowIndex) ?? sheet.createRow(rowIndex);
        // 0xFF is the "use default" marker — only set explicit height for other values
        if (heightTwips > 0 && heightTwips != 0xFF)
        {
            row.setHeight((float)heightTwips / 20.0f);
        }
        if (hidden) row.setHidden(true);
    }

    private static void ReadColInfo(HSSFSheet sheet, ReadOnlySpan<byte> data)
    {
        if (data.Length < 10) return;
        var firstCol = BinaryPrimitives.ReadUInt16LittleEndian(data);
        var lastCol = BinaryPrimitives.ReadUInt16LittleEndian(data.Slice(2));
        var colWidth = BinaryPrimitives.ReadUInt16LittleEndian(data.Slice(4));
        var options = BinaryPrimitives.ReadUInt16LittleEndian(data.Slice(8));
        var hidden = (options & 0x01) != 0;

        for (var col = firstCol; col <= lastCol; col++)
        {
            if (colWidth != 2275) sheet.setColumnWidth(col, colWidth); // 2275 = default
            if (hidden) sheet.setColumnHidden(col, true);
        }
    }

    private static void ReadMergeCells(HSSFSheet sheet, ReadOnlySpan<byte> data)
    {
        if (data.Length < 2) return;
        var count = BinaryPrimitives.ReadUInt16LittleEndian(data);
        var pos = 2;
        for (var i = 0; i < count; i++)
        {
            if (pos + 8 > data.Length) break;
            var firstRow = BinaryPrimitives.ReadUInt16LittleEndian(data.Slice(pos));
            var lastRow = BinaryPrimitives.ReadUInt16LittleEndian(data.Slice(pos + 2));
            var firstCol = BinaryPrimitives.ReadUInt16LittleEndian(data.Slice(pos + 4));
            var lastCol = BinaryPrimitives.ReadUInt16LittleEndian(data.Slice(pos + 6));
            sheet.addMergedRegion(new SS.Util.CellRangeAddress(firstRow, lastRow, firstCol, lastCol));
            pos += 8;
        }
    }

    private static void ReadPaneRecord(HSSFSheet sheet, ReadOnlySpan<byte> data)
    {
        if (data.Length < 10) return;
        var xSplit = BinaryPrimitives.ReadUInt16LittleEndian(data);
        var ySplit = BinaryPrimitives.ReadUInt16LittleEndian(data.Slice(2));
        // In freeze pane mode, x/y are column/row counts directly
        if (xSplit > 0 || ySplit > 0)
            sheet.createFreezePane(xSplit, ySplit);
    }

    private static void ReadLabelSst(HSSFWorkbook workbook, HSSFSheet sheet, ReadOnlySpan<byte> data, IReadOnlyList<string> sharedStrings)
    {
        if (data.Length < 10) return;
        var cell = GetOrCreateCell(sheet, ReadUInt16(data, 0), ReadUInt16(data, 2));
        SetCellStyleFromXfIndex(workbook, cell, ReadUInt16(data, 4));
        var index = BinaryPrimitives.ReadUInt32LittleEndian(data.Slice(6));
        cell.setCellValue(index < sharedStrings.Count ? sharedStrings[(int)index] : string.Empty);
    }

    private static void ReadLabel(HSSFWorkbook workbook, HSSFSheet sheet, ReadOnlySpan<byte> data)
    {
        if (data.Length < 8) return;
        var cell = GetOrCreateCell(sheet, ReadUInt16(data, 0), ReadUInt16(data, 2));
        SetCellStyleFromXfIndex(workbook, cell, ReadUInt16(data, 4));
        var pos = 6;
        cell.setCellValue(ReadUnicodeString(data, ref pos, shortLength: false));
    }

    private static void ReadNumber(HSSFWorkbook workbook, HSSFSheet sheet, ReadOnlySpan<byte> data)
    {
        if (data.Length < 14) return;
        var cell = GetOrCreateCell(sheet, ReadUInt16(data, 0), ReadUInt16(data, 2));
        SetCellStyleFromXfIndex(workbook, cell, ReadUInt16(data, 4));
        cell.setCellValue(BitConverter.Int64BitsToDouble(BinaryPrimitives.ReadInt64LittleEndian(data.Slice(6))));
    }

    private static void ReadRk(HSSFWorkbook workbook, HSSFSheet sheet, ReadOnlySpan<byte> data)
    {
        if (data.Length < 10) return;
        var cell = GetOrCreateCell(sheet, ReadUInt16(data, 0), ReadUInt16(data, 2));
        SetCellStyleFromXfIndex(workbook, cell, ReadUInt16(data, 4));
        cell.setCellValue(DecodeRk(BinaryPrimitives.ReadInt32LittleEndian(data.Slice(6))));
    }

    private static void ReadBoolErr(HSSFWorkbook workbook, HSSFSheet sheet, ReadOnlySpan<byte> data)
    {
        if (data.Length < 8) return;
        var cell = GetOrCreateCell(sheet, ReadUInt16(data, 0), ReadUInt16(data, 2));
        SetCellStyleFromXfIndex(workbook, cell, ReadUInt16(data, 4));
        if (data[7] == 1) cell.SetError(data[6]);
        else cell.setCellValue(data[6] != 0);
    }

    private static void ReadBlank(HSSFWorkbook workbook, HSSFSheet sheet, ReadOnlySpan<byte> data)
    {
        if (data.Length < 6) return;
        var cell = GetOrCreateCell(sheet, ReadUInt16(data, 0), ReadUInt16(data, 2));
        SetCellStyleFromXfIndex(workbook, cell, ReadUInt16(data, 4));
        cell.SetBlank();
    }

    private static HSSFCell GetOrCreateCell(HSSFSheet sheet, int rowIndex, int columnIndex)
    {
        var row = sheet.getRow(rowIndex) ?? sheet.createRow(rowIndex);
        return row.getCell(columnIndex) ?? row.createCell(columnIndex);
    }

    private static List<Record> ReadRecords(byte[] stream)
    {
        var records = new List<Record>();
        var offset = 0;
        while (offset + 4 <= stream.Length)
        {
            var sid = BinaryPrimitives.ReadUInt16LittleEndian(stream.AsSpan(offset));
            var length = BinaryPrimitives.ReadUInt16LittleEndian(stream.AsSpan(offset + 2));
            if (offset + 4 + length > stream.Length)
            {
                break;
            }

            var data = stream.AsSpan(offset + 4, length).ToArray();
            records.Add(new Record(sid, checked((uint)offset), data));
            offset += 4 + length;
        }

        return records;
    }

    private static BoundSheet ReadBoundSheet(ReadOnlySpan<byte> data)
    {
        if (data.Length < 8)
        {
            return new BoundSheet(0, "Sheet");
        }

        var offset = BinaryPrimitives.ReadUInt32LittleEndian(data);
        var pos = 6;
        return new BoundSheet(offset, ReadUnicodeString(data, ref pos, shortLength: true));
    }

    private static List<string> ReadSst(ReadOnlySpan<byte> data)
    {
        var strings = new List<string>();
        if (data.Length < 8)
        {
            return strings;
        }

        var uniqueCount = BinaryPrimitives.ReadUInt32LittleEndian(data.Slice(4));
        var pos = 8;
        for (var i = 0; i < uniqueCount && pos < data.Length; i++)
        {
            strings.Add(ReadUnicodeString(data, ref pos, shortLength: false));
        }

        return strings;
    }

    private static string ReadUnicodeString(ReadOnlySpan<byte> data, ref int pos, bool shortLength)
    {
        if (pos >= data.Length)
        {
            return string.Empty;
        }

        var charCount = shortLength ? data[pos++] : ReadUInt16AndAdvance(data, ref pos);
        if (pos >= data.Length)
        {
            return string.Empty;
        }

        var options = data[pos++];
        var hasExt = (options & 0x04) != 0;
        var hasRich = (options & 0x08) != 0;
        var isUtf16 = (options & 0x01) != 0;
        var richRuns = hasRich && pos + 2 <= data.Length ? ReadUInt16AndAdvance(data, ref pos) : 0;
        var extSize = hasExt && pos + 4 <= data.Length ? ReadUInt32AndAdvance(data, ref pos) : 0;
        var byteCount = charCount * (isUtf16 ? 2 : 1);
        if (pos + byteCount > data.Length)
        {
            byteCount = Math.Max(0, data.Length - pos);
        }

        var value = isUtf16
            ? Encoding.Unicode.GetString(data.Slice(pos, byteCount).ToArray())
            : Encoding.GetEncoding("ISO-8859-1").GetString(data.Slice(pos, byteCount).ToArray());
        pos += byteCount;
        pos += Math.Min(richRuns * 4, Math.Max(0, data.Length - pos));
        pos += (int)Math.Min(extSize, (uint)Math.Max(0, data.Length - pos));
        return value;
    }

    private static Dictionary<string, int> BuildSharedStrings(
        IReadOnlyList<HSSFSheet> sheets,
        IEnumerable<string>? existingStrings = null)
    {
        var strings = new Dictionary<string, int>(StringComparer.Ordinal);
        foreach (var value in existingStrings ?? Array.Empty<string>())
        {
            if (!strings.ContainsKey(value))
            {
                strings[value] = strings.Count;
            }
        }

        foreach (var cell in sheets.SelectMany(sheet => sheet.Rows).SelectMany(row => row.Cells))
        {
            if (cell.getCellType() != CellType.String)
            {
                continue;
            }

            var value = cell.getStringCellValue();
            if (!strings.ContainsKey(value))
            {
                strings[value] = strings.Count;
            }
        }

        return strings;
    }

    private static void WriteSst(Stream stream, IReadOnlyDictionary<string, int> strings)
    {
        WriteRecord(stream, Sst, payload =>
        {
            Span<byte> header = stackalloc byte[8];
            BinaryPrimitives.WriteUInt32LittleEndian(header, (uint)strings.Count);
            BinaryPrimitives.WriteUInt32LittleEndian(((Span<byte>)header).Slice(4), (uint)strings.Count);
            payload.Write(header.ToArray(), 0, header.Length);

            foreach (var value in strings.OrderBy(kv => kv.Value).Select(kv => kv.Key))
            {
                WriteUnicodeString(payload, value, shortLength: false);
            }
        });
    }

    private static void WriteSheet(Stream stream, HSSFSheet sheet, IReadOnlyDictionary<string, int> strings)
    {
        WriteBof(stream, 0x0010);
        WriteColInfoRecords(stream, sheet);
        WriteDimensions(stream, sheet);
        // Write row blocks: each row's ROW record followed by its cells
        WriteRowBlocksAndCells(stream, sheet, strings);
        WriteMergeCellsRecord(stream, sheet);
        WriteWindow2(stream);
        if (sheet.FreezeColSplit > 0 || sheet.FreezeRowSplit > 0)
            WritePaneRecord(stream, sheet);
        WriteSelection(stream);
        WriteRecord(stream, Eof, _ => { });
    }

    private static void WriteRowBlocksAndCells(Stream stream, HSSFSheet sheet, IReadOnlyDictionary<string, int> strings)
    {
        foreach (var row in sheet.Rows.OrderBy(r => r.getRowNum()))
        {
            // Write RowRecord
            WriteRowRecord(stream, row);
            // Write cells for this row
            foreach (var cell in row.Cells.OrderBy(c => c.getColumnIndex()))
            {
                switch (cell.getCellType())
                {
                    case CellType.String:
                        WriteCellPrefixRecord(stream, LabelSst, cell, payload =>
                        {
                            Span<byte> index = stackalloc byte[4];
                            BinaryPrimitives.WriteUInt32LittleEndian(index, (uint)strings[cell.getStringCellValue()]);
                            payload.Write(index.ToArray(), 0, index.Length);
                        });
                        break;
                    case CellType.Numeric:
                        WriteCellPrefixRecord(stream, Number, cell, payload =>
                        {
                            Span<byte> value = stackalloc byte[8];
                            BinaryPrimitives.WriteInt64LittleEndian(value, BitConverter.DoubleToInt64Bits(cell.getNumericCellValue()));
                            payload.Write(value.ToArray(), 0, value.Length);
                        });
                        break;
                    case CellType.Boolean:
                        WriteCellPrefixRecord(stream, BoolErr, cell, payload =>
                        {
                            payload.WriteByte(cell.getBooleanCellValue() ? (byte)1 : (byte)0);
                            payload.WriteByte(0);
                        });
                        break;
                    case CellType.Error:
                        WriteCellPrefixRecord(stream, BoolErr, cell, payload =>
                        {
                            payload.WriteByte(cell.GetErrorByte());
                            payload.WriteByte(1);
                        });
                        break;
                    case CellType.Blank:
                        WriteCellPrefixRecord(stream, Blank, cell, _ => { });
                        break;
                }
            }
        }
    }

    private static void WriteRowRecord(Stream stream, HSSFRow row)
    {
        WriteRecord(stream, RowRecord, payload =>
        {
            Span<byte> buf = stackalloc byte[16];
            buf.Clear();
            BinaryPrimitives.WriteUInt16LittleEndian(buf, (ushort)row.getRowNum());
            var rowCells = row.Cells.ToList();
            BinaryPrimitives.WriteUInt16LittleEndian(buf.Slice(2), rowCells.Count == 0 ? (ushort)0 : (ushort)rowCells.Min(c => c.getColumnIndex()));
            BinaryPrimitives.WriteUInt16LittleEndian(buf.Slice(4), rowCells.Count == 0 ? (ushort)0 : (ushort)(rowCells.Max(c => c.getColumnIndex()) + 1));
            // Height: 0xFF = use default; custom height in twips otherwise
            var height = row.HeightTwips > 0 ? (ushort)row.HeightTwips : (ushort)0xFF;
            BinaryPrimitives.WriteUInt16LittleEndian(buf.Slice(6), height);
            // field_5_optimize=0, field_6_reserved=0 (bytes 8-11 are zero)
            ushort optionFlags = 0x0100; // OPTION_BITS_ALWAYS_SET
            if (row.Hidden) optionFlags |= 0x0020; // zeroHeight bit
            BinaryPrimitives.WriteUInt16LittleEndian(buf.Slice(12), optionFlags);
            BinaryPrimitives.WriteUInt16LittleEndian(buf.Slice(14), 0x000F); // field_8
            payload.Write(buf.ToArray(), 0, buf.Length);
        });
    }

    private static void WriteColInfoRecords(Stream stream, HSSFSheet sheet)
    {
        // Collect the range of columns that might need ColInfo records
        int maxCol = -1;
        foreach (var col in sheet.ExplicitColumnWidthIndices) maxCol = Math.Max(maxCol, col);
        foreach (var col in sheet.HiddenColumns) maxCol = Math.Max(maxCol, col);

        if (maxCol == -1) return;

        int i = 0;
        while (i <= maxCol)
        {
            var width = sheet.getColumnWidth(i);
            var hidden = sheet.isColumnHidden(i);
            var biffWidth = width > 0 ? (ushort)width : (ushort)2275; // 2275 = default

            // Skip default columns if they are not hidden
            if (biffWidth == 2275 && !hidden)
            {
                i++;
                continue;
            }

            int first = i;
            int last = i;

            // Group contiguous columns with same effective BIFF settings
            while (i + 1 <= maxCol)
            {
                var nextWidth = sheet.getColumnWidth(i + 1);
                var nextHidden = sheet.isColumnHidden(i + 1);
                var nextBiffWidth = nextWidth > 0 ? (ushort)nextWidth : (ushort)2275;

                if (nextBiffWidth == biffWidth && nextHidden == hidden)
                {
                    last = ++i;
                }
                else
                {
                    break;
                }
            }
            i++;

            WriteRecord(stream, ColInfo, payload =>
            {
                Span<byte> buf = stackalloc byte[12];
                buf.Clear();
                BinaryPrimitives.WriteUInt16LittleEndian(buf, (ushort)first);
                BinaryPrimitives.WriteUInt16LittleEndian(buf.Slice(2), (ushort)last);
                BinaryPrimitives.WriteUInt16LittleEndian(buf.Slice(4), biffWidth);
                BinaryPrimitives.WriteUInt16LittleEndian(buf.Slice(6), 0x000F); // XF index
                ushort options = 0x0002; // flag bit
                if (hidden) options |= 0x0001;
                BinaryPrimitives.WriteUInt16LittleEndian(buf.Slice(8), options);
                BinaryPrimitives.WriteUInt16LittleEndian(buf.Slice(10), 0x0002); // reserved
                payload.Write(buf.ToArray(), 0, buf.Length);
            });
        }
    }

    private static void WriteMergeCellsRecord(Stream stream, HSSFSheet sheet)
    {
        var regions = sheet.getMergedRegions();
        if (regions.Count == 0) return;

        WriteRecord(stream, MergeCellsRecord, payload =>
        {
            Span<byte> countBuf = stackalloc byte[2];
            BinaryPrimitives.WriteUInt16LittleEndian(countBuf, (ushort)regions.Count);
            payload.Write(countBuf.ToArray(), 0, countBuf.Length);
            foreach (var region in regions)
            {
                Span<byte> regionBuf = stackalloc byte[8];
                BinaryPrimitives.WriteUInt16LittleEndian(regionBuf, (ushort)region.FirstRow);
                BinaryPrimitives.WriteUInt16LittleEndian(regionBuf.Slice(2), (ushort)region.LastRow);
                BinaryPrimitives.WriteUInt16LittleEndian(regionBuf.Slice(4), (ushort)region.FirstCol);
                BinaryPrimitives.WriteUInt16LittleEndian(regionBuf.Slice(6), (ushort)region.LastCol);
                payload.Write(regionBuf.ToArray(), 0, regionBuf.Length);
            }
        });
    }

    private static void WritePaneRecord(Stream stream, HSSFSheet sheet)
    {
        WriteRecord(stream, PaneRecord, payload =>
        {
            Span<byte> buf = stackalloc byte[10];
            buf.Clear();
            BinaryPrimitives.WriteUInt16LittleEndian(buf, (ushort)sheet.FreezeColSplit);
            BinaryPrimitives.WriteUInt16LittleEndian(buf.Slice(2), (ushort)sheet.FreezeRowSplit);
            BinaryPrimitives.WriteUInt16LittleEndian(buf.Slice(4), (ushort)sheet.FreezeRowSplit); // topRow
            BinaryPrimitives.WriteUInt16LittleEndian(buf.Slice(6), (ushort)sheet.FreezeColSplit); // leftCol
            buf[8] = 0; // active pane = 0 (bottom-right)
            payload.Write(buf.ToArray(), 0, buf.Length);
        });
    }

    private static void WriteSheetCells(Stream stream, HSSFSheet sheet, IReadOnlyDictionary<string, int> strings)
    {
        foreach (var cell in sheet.Rows.OrderBy(row => row.getRowNum()).SelectMany(row => row.Cells.OrderBy(cell => cell.getColumnIndex())))
        {
            switch (cell.getCellType())
            {
                case CellType.String:
                    WriteCellPrefixRecord(stream, LabelSst, cell, payload =>
                    {
                        Span<byte> index = stackalloc byte[4];
                        BinaryPrimitives.WriteUInt32LittleEndian(index, (uint)strings[cell.getStringCellValue()]);
                        payload.Write(index.ToArray(), 0, index.Length);
                    });
                    break;
                case CellType.Numeric:
                    WriteCellPrefixRecord(stream, Number, cell, payload =>
                    {
                        Span<byte> value = stackalloc byte[8];
                        BinaryPrimitives.WriteInt64LittleEndian(value, BitConverter.DoubleToInt64Bits(cell.getNumericCellValue()));
                        payload.Write(value.ToArray(), 0, value.Length);
                    });
                    break;
                case CellType.Boolean:
                    WriteCellPrefixRecord(stream, BoolErr, cell, payload =>
                    {
                        payload.WriteByte(cell.getBooleanCellValue() ? (byte)1 : (byte)0);
                        payload.WriteByte(0);
                    });
                    break;
                case CellType.Error:
                    WriteCellPrefixRecord(stream, BoolErr, cell, payload =>
                    {
                        payload.WriteByte(cell.GetErrorByte());
                        payload.WriteByte(1);
                    });
                    break;
                case CellType.Blank:
                    WriteCellPrefixRecord(stream, Blank, cell, _ => { });
                    break;
            }
        }
    }

    private static void WriteSheetPreservingRecords(
        Stream stream,
        HSSFSheet sheet,
        IReadOnlyDictionary<string, int> strings,
        IReadOnlyList<Record> originalRecords)
    {
        if (originalRecords.Count == 0)
        {
            WriteSheet(stream, sheet, strings);
            return;
        }

        var wroteDimensionsAndCells = false;
        foreach (var record in originalRecords)
        {
            switch (record.Sid)
            {
                case Dimensions:
                    WriteDimensions(stream, sheet);
                    WriteRowBlocksAndCells(stream, sheet, strings);
                    wroteDimensionsAndCells = true;
                    break;
                case RowRecord:
                case LabelSst:
                case Label:
                case Number:
                case Rk:
                case BoolErr:
                case Blank:
                    break;
                case Eof:
                    if (!wroteDimensionsAndCells)
                    {
                        WriteDimensions(stream, sheet);
                        WriteRowBlocksAndCells(stream, sheet, strings);
                    }

                    WriteRawRecord(stream, record);
                    break;
                default:
                    WriteRawRecord(stream, record);
                    break;
            }
        }
    }

    private static void WriteWindow2(Stream stream)
    {
        WriteRecord(stream, Window2, payload =>
        {
            Span<byte> buffer = stackalloc byte[18];
            buffer.Clear();
            BinaryPrimitives.WriteUInt16LittleEndian(buffer, 0x06B6);
            BinaryPrimitives.WriteUInt32LittleEndian(((Span<byte>)buffer).Slice(6), 0x40);
            payload.Write(buffer.ToArray(), 0, buffer.Length);
        });
    }

    private static void WriteSelection(Stream stream)
    {
        WriteRecord(stream, Selection, payload =>
        {
            Span<byte> buffer = stackalloc byte[15];
            buffer.Clear();
            buffer[0] = 0x03; // pane
            BinaryPrimitives.WriteUInt16LittleEndian(buffer.Slice(7), 1); // one selected range
            // Ref8U for A1:A1: rwFirst/rwLast/colFirst/colLast are zero.
            payload.Write(buffer.ToArray(), 0, buffer.Length);
        });
    }

    private static void WriteDimensions(Stream stream, HSSFSheet sheet)
    {
        var rows = sheet.Rows.ToArray();
        var firstRow = rows.Length == 0 ? 0 : rows.Min(row => row.getRowNum());
        var lastRow = rows.Length == 0 ? 0 : rows.Max(row => row.getRowNum()) + 1;
        var cells = rows.SelectMany(row => row.Cells).ToArray();
        var firstCol = cells.Length == 0 ? 0 : cells.Min(cell => cell.getColumnIndex());
        var lastCol = cells.Length == 0 ? 0 : cells.Max(cell => cell.getColumnIndex()) + 1;

        WriteRecord(stream, Dimensions, payload =>
        {
            Span<byte> buffer = stackalloc byte[14];
            buffer.Clear();
            BinaryPrimitives.WriteUInt32LittleEndian(buffer, (uint)firstRow);
            BinaryPrimitives.WriteUInt32LittleEndian(((Span<byte>)buffer).Slice(4), (uint)lastRow);
            BinaryPrimitives.WriteUInt16LittleEndian(((Span<byte>)buffer).Slice(8), (ushort)firstCol);
            BinaryPrimitives.WriteUInt16LittleEndian(((Span<byte>)buffer).Slice(10), (ushort)lastCol);
            payload.Write(buffer.ToArray(), 0, buffer.Length);
        });
    }

    private static void WriteCellPrefixRecord(Stream stream, ushort sid, HSSFCell cell, Action<Stream> writeRemainder)
    {
        WriteRecord(stream, sid, payload =>
        {
            Span<byte> prefix = stackalloc byte[6];
            BinaryPrimitives.WriteUInt16LittleEndian(prefix, (ushort)cell.getRowIndex());
            BinaryPrimitives.WriteUInt16LittleEndian(((Span<byte>)prefix).Slice(2), (ushort)cell.getColumnIndex());
            BinaryPrimitives.WriteUInt16LittleEndian(((Span<byte>)prefix).Slice(4), cell.GetXfIndex());
            payload.Write(prefix.ToArray(), 0, prefix.Length);
            writeRemainder(payload);
        });
    }

    private static void WriteBoundSheet(Stream stream, string name, uint offset)
    {
        WriteRecord(stream, BoundSheet8, payload =>
        {
            Span<byte> buffer = stackalloc byte[6];
            buffer.Clear();
            BinaryPrimitives.WriteUInt32LittleEndian(buffer, offset);
            payload.Write(buffer.ToArray(), 0, buffer.Length);
            WriteUnicodeString(payload, name, shortLength: true);
        });
    }

    private static void WriteBof(Stream stream, ushort type)
    {
        WriteRecord(stream, Bof, payload =>
        {
            Span<byte> buffer = stackalloc byte[16];
            BinaryPrimitives.WriteUInt16LittleEndian(buffer, 0x0600);
            BinaryPrimitives.WriteUInt16LittleEndian(((Span<byte>)buffer).Slice(2), type);
            BinaryPrimitives.WriteUInt16LittleEndian(((Span<byte>)buffer).Slice(4), 0x0DBB);
            BinaryPrimitives.WriteUInt16LittleEndian(((Span<byte>)buffer).Slice(6), 0x07CC);
            BinaryPrimitives.WriteUInt32LittleEndian(((Span<byte>)buffer).Slice(8), 0x00000041);
            BinaryPrimitives.WriteUInt32LittleEndian(((Span<byte>)buffer).Slice(12), 0x00000006);
            payload.Write(buffer.ToArray(), 0, buffer.Length);
        });
    }

    private static void WriteRecord(Stream stream, ushort sid, Action<Stream> writePayload)
    {
        using var payload = new MemoryStream();
        writePayload(payload);
        var data = payload.ToArray();
        if (data.Length > ushort.MaxValue)
        {
            throw new InvalidDataException("BIFF record exceeds the maximum record size.");
        }

        Span<byte> header = stackalloc byte[4];
        BinaryPrimitives.WriteUInt16LittleEndian(header, sid);
        BinaryPrimitives.WriteUInt16LittleEndian(((Span<byte>)header).Slice(2), (ushort)data.Length);
        stream.Write(header.ToArray(), 0, header.Length);
        stream.Write(data, 0, data.Length);
    }

    private static void WriteRawRecord(Stream stream, Record record)
    {
        Span<byte> header = stackalloc byte[4];
        BinaryPrimitives.WriteUInt16LittleEndian(header, record.Sid);
        BinaryPrimitives.WriteUInt16LittleEndian(((Span<byte>)header).Slice(2), (ushort)record.Data.Length);
        stream.Write(header.ToArray(), 0, header.Length);
        stream.Write(record.Data, 0, record.Data.Length);
    }

    private static IReadOnlyList<Record> GetSheetRecords(IReadOnlyList<Record> records, uint offset)
    {
        var start = 0;
        while (start < records.Count && records[start].Offset < offset)
        {
            start++;
        }

        var sheetRecords = new List<Record>();
        for (var i = start; i < records.Count; i++)
        {
            var record = records[i];
            if (record.Offset > offset && record.Sid == Bof && i != start)
            {
                break;
            }

            sheetRecords.Add(record);
            if (record.Offset >= offset && record.Sid == Eof)
            {
                break;
            }
        }

        return sheetRecords;
    }

    private static void WriteUnicodeString(Stream stream, string value, bool shortLength)
    {
        var compressed = value.All(ch => ch <= 0x00FF);
        if (shortLength)
        {
            stream.WriteByte((byte)Math.Min(value.Length, byte.MaxValue));
        }
        else
        {
            Span<byte> length = stackalloc byte[2];
            BinaryPrimitives.WriteUInt16LittleEndian(length, (ushort)Math.Min(value.Length, ushort.MaxValue));
            stream.Write(length.ToArray(), 0, length.Length);
        }

        stream.WriteByte(compressed ? (byte)0 : (byte)1);
        stream.Write(compressed ? Encoding.GetEncoding("ISO-8859-1").GetBytes(value) : Encoding.Unicode.GetBytes(value), 0, (compressed ? Encoding.GetEncoding("ISO-8859-1").GetBytes(value) : Encoding.Unicode.GetBytes(value)).Length);
    }

    private static void PatchUInt32(Stream stream, long position, uint value)
    {
        var current = stream.Position;
        stream.Position = position;
        Span<byte> buffer = stackalloc byte[4];
        BinaryPrimitives.WriteUInt32LittleEndian(buffer, value);
        stream.Write(buffer.ToArray(), 0, buffer.Length);
        stream.Position = current;
    }

    private static double DecodeRk(int rk)
    {
        double value;
        if ((rk & 0x02) != 0)
        {
            value = rk >> 2;
        }
        else
        {
            Span<byte> bytes = stackalloc byte[8];
            BinaryPrimitives.WriteInt32LittleEndian(((Span<byte>)bytes).Slice(4), rk & unchecked((int)0xFFFFFFFC));
            value = BitConverter.Int64BitsToDouble(BinaryPrimitives.ReadInt64LittleEndian(bytes));
        }

        return (rk & 0x01) != 0 ? value / 100.0 : value;
    }

    private static ushort ReadUInt16(ReadOnlySpan<byte> data, int offset) =>
        BinaryPrimitives.ReadUInt16LittleEndian(data.Slice(offset));

    private static ushort ReadUInt16AndAdvance(ReadOnlySpan<byte> data, ref int pos)
    {
        var value = BinaryPrimitives.ReadUInt16LittleEndian(data.Slice(pos));
        pos += 2;
        return value;
    }

    private static uint ReadUInt32AndAdvance(ReadOnlySpan<byte> data, ref int pos)
    {
        var value = BinaryPrimitives.ReadUInt32LittleEndian(data.Slice(pos));
        pos += 4;
        return value;
    }

    private sealed record BoundSheet(uint Offset, string Name);

    private sealed record Record(ushort Sid, uint Offset, byte[] Data);
}
