using System.Buffers.Binary;
using System.Text;
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

    public static void ReadWorkbook(byte[] workbookStream, HSSFWorkbook workbook)
    {
        var records = ReadRecords(workbookStream);
        var sheets = new List<BoundSheet>();
        var sharedStrings = new List<string>();

        foreach (var record in records)
        {
            switch (record.Sid)
            {
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

    public static byte[] WriteWorkbook(IReadOnlyList<HSSFSheet> sheets)
    {
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

    private static void ReadSheet(IReadOnlyList<Record> records, uint offset, HSSFSheet sheet, IReadOnlyList<string> sharedStrings)
    {
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
                case LabelSst:
                    ReadLabelSst(sheet, record.Data, sharedStrings);
                    break;
                case Label:
                    ReadLabel(sheet, record.Data);
                    break;
                case Number:
                    ReadNumber(sheet, record.Data);
                    break;
                case Rk:
                    ReadRk(sheet, record.Data);
                    break;
                case BoolErr:
                    ReadBoolErr(sheet, record.Data);
                    break;
                case Blank:
                    ReadBlank(sheet, record.Data);
                    break;
            }
        }
    }

    private static void ReadLabelSst(HSSFSheet sheet, ReadOnlySpan<byte> data, IReadOnlyList<string> sharedStrings)
    {
        if (data.Length < 10)
        {
            return;
        }

        var cell = GetOrCreateCell(sheet, ReadUInt16(data, 0), ReadUInt16(data, 2));
        var index = BinaryPrimitives.ReadUInt32LittleEndian(data.Slice(6));
        cell.setCellValue(index < sharedStrings.Count ? sharedStrings[(int)index] : string.Empty);
    }

    private static void ReadLabel(HSSFSheet sheet, ReadOnlySpan<byte> data)
    {
        if (data.Length < 8)
        {
            return;
        }

        var cell = GetOrCreateCell(sheet, ReadUInt16(data, 0), ReadUInt16(data, 2));
        var pos = 6;
        cell.setCellValue(ReadUnicodeString(data, ref pos, shortLength: false));
    }

    private static void ReadNumber(HSSFSheet sheet, ReadOnlySpan<byte> data)
    {
        if (data.Length < 14)
        {
            return;
        }

        var cell = GetOrCreateCell(sheet, ReadUInt16(data, 0), ReadUInt16(data, 2));
        cell.setCellValue(BitConverter.Int64BitsToDouble(BinaryPrimitives.ReadInt64LittleEndian(data.Slice(6))));
    }

    private static void ReadRk(HSSFSheet sheet, ReadOnlySpan<byte> data)
    {
        if (data.Length < 10)
        {
            return;
        }

        var cell = GetOrCreateCell(sheet, ReadUInt16(data, 0), ReadUInt16(data, 2));
        cell.setCellValue(DecodeRk(BinaryPrimitives.ReadInt32LittleEndian(data.Slice(6))));
    }

    private static void ReadBoolErr(HSSFSheet sheet, ReadOnlySpan<byte> data)
    {
        if (data.Length < 8)
        {
            return;
        }

        var cell = GetOrCreateCell(sheet, ReadUInt16(data, 0), ReadUInt16(data, 2));
        if (data[7] == 1)
        {
            cell.SetError(data[6]);
        }
        else
        {
            cell.setCellValue(data[6] != 0);
        }
    }

    private static void ReadBlank(HSSFSheet sheet, ReadOnlySpan<byte> data)
    {
        if (data.Length < 6)
        {
            return;
        }

        GetOrCreateCell(sheet, ReadUInt16(data, 0), ReadUInt16(data, 2)).SetBlank();
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

    private static Dictionary<string, int> BuildSharedStrings(IReadOnlyList<HSSFSheet> sheets)
    {
        var strings = new Dictionary<string, int>(StringComparer.Ordinal);
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
        WriteDimensions(stream, sheet);

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
                case CellType.Blank:
                    WriteCellPrefixRecord(stream, Blank, cell, _ => { });
                    break;
            }
        }

        WriteWindow2(stream);
        WriteSelection(stream);
        WriteRecord(stream, Eof, _ => { });
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
            Span<byte> buffer = stackalloc byte[9];
            buffer.Clear();
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
            BinaryPrimitives.WriteUInt16LittleEndian(((Span<byte>)prefix).Slice(4), 0);
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
