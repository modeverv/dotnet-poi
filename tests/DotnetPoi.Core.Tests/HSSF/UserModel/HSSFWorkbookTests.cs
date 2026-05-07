using DotnetPoi.HSSF.UserModel;
using DotnetPoi.POIFS.Crypt;
using DotnetPoi.SS.UserModel;
using System.Buffers.Binary;
using Xunit;

namespace DotnetPoi.HSSF.Tests.UserModel;

public sealed class HSSFWorkbookTests
{
    public static IEnumerable<object[]> Phase12RepresentativePoiFixtures()
    {
        yield return new object[] { "empty.xls", 3 };
        yield return new object[] { "Simple.xls", 3 };
        yield return new object[] { "SimpleMultiCell.xls", 3 };
        yield return new object[] { "SampleSS.xls", 3 };
        yield return new object[] { "WORKBOOK_in_capitals.xls", 1 };
        yield return new object[] { "BOOK_in_capitals.xls", 1 };
        yield return new object[] { "chinese-provinces.xls", 1 };
        yield return new object[] { "DateFormats.xls", 3 };
        yield return new object[] { "SimpleWithStyling.xls", 3 };
        yield return new object[] { "WithExtendedStyles.xls", 3 };
        yield return new object[] { "55341_CellStyleBorder.xls", 3 };
        yield return new object[] { "SimpleWithPrintArea.xls", 3 };
        yield return new object[] { "RepeatingRowsCols.xls", 4 };
        yield return new object[] { "SimpleWithFormula.xls", 3 };
        yield return new object[] { "ex47747-sharedFormula.xls", 1 };
        yield return new object[] { "WithHyperlink.xls", 3 };
        yield return new object[] { "comments.xls", 3 };
        yield return new object[] { "drawings.xls", 8 };
        yield return new object[] { "SimpleWithImages.xls", 3 };
        yield return new object[] { "SimpleMacro.xls", 3 };
    }

    [Fact]
    public void Write_StringNumberBooleanCells_RoundTrips()
    {
        using var workbook = new HSSFWorkbook();
        var sheet = workbook.createSheet("Data");
        var row = sheet.createRow(0);
        row.createCell(0).setCellValue("alpha");
        row.createCell(1).setCellValue(42.5);
        row.createCell(2).setCellValue(true);

        using var stream = new MemoryStream();
        workbook.write(stream);
        stream.Position = 0;

        using var read = new HSSFWorkbook(stream);

        Assert.Equal(1, read.getNumberOfSheets());
        var readRow = read.getSheetAt(0).getRow(0);
        Assert.NotNull(readRow);
        Assert.Equal("alpha", readRow!.getCell(0)!.getStringCellValue());
        Assert.Equal(42.5, readRow.getCell(1)!.getNumericCellValue());
        Assert.True(readRow.getCell(2)!.getBooleanCellValue());
    }

    [Fact]
    public void Read_PoiSampleXls_LoadsWorkbookStreamAndBasicCells()
    {
        var sample = FindRepoRoot().Combine("poi/test-data/spreadsheet/SimpleMultiCell.xls");
        using var input = sample.OpenRead();

        using var workbook = new HSSFWorkbook(input);

        Assert.True(workbook.getNumberOfSheets() > 0);
        var sheet = workbook.getSheetAt(0);
        Assert.True(sheet.getLastRowNum() >= 0);
        Assert.Contains(Enumerable.Range(0, sheet.getLastRowNum() + 1), rowIndex =>
        {
            var row = sheet.getRow(rowIndex);
            return row is not null && Enumerable.Range(0, Math.Max(row.getLastCellNum(), (short)0))
                .Select(row.getCell)
                .Any(cell => cell is not null && cell.getCellType() is CellType.String or CellType.Numeric or CellType.Boolean);
        });
    }

    [Theory]
    [MemberData(nameof(Phase12RepresentativePoiFixtures))]
    public void Read_Phase12RepresentativePoiFixtures_LoadsWorkbookStream(string fileName, int expectedSheetCount)
    {
        var sample = FindRepoRoot().Combine(Path.Combine("poi/test-data/spreadsheet", fileName));
        using var input = sample.OpenRead();

        using var workbook = new HSSFWorkbook(input);

        Assert.Equal(expectedSheetCount, workbook.getNumberOfSheets());
    }

    [Fact]
    public void Read_BookUppercasePoiFixture_LoadsWorkbookStream()
    {
        var sample = FindRepoRoot().Combine("poi/test-data/spreadsheet/BOOK_in_capitals.xls");
        using var input = sample.OpenRead();

        using var workbook = new HSSFWorkbook(input);

        Assert.Equal(1, workbook.getNumberOfSheets());
    }

    [Fact]
    public void Write_LoadedPoiFixture_PreservesNonWorkbookOleStreams()
    {
        var sample = FindRepoRoot().Combine("poi/test-data/spreadsheet/empty.xls");
        Dictionary<string, byte[]> originalStreams;
        using (var originalInput = sample.OpenRead())
        {
            originalStreams = CompoundFile.ReadStreams(originalInput);
        }

        using var workbookInput = sample.OpenRead();
        using var workbook = new HSSFWorkbook(workbookInput);

        using var output = new MemoryStream();
        workbook.write(output);
        output.Position = 0;
        var writtenStreams = CompoundFile.ReadStreams(output);

        foreach (var streamName in originalStreams.Keys.Where(name => name != "Workbook"))
        {
            Assert.True(writtenStreams.ContainsKey(streamName), $"Missing preserved stream '{streamName}'.");
            Assert.Equal(originalStreams[streamName], writtenStreams[streamName]);
        }
    }

    [Fact]
    public void Write_LoadedMacroPoiFixture_PreservesNestedOleStorageStreams()
    {
        var sample = FindRepoRoot().Combine("poi/test-data/spreadsheet/SimpleMacro.xls");
        Dictionary<string, byte[]> originalStreams;
        using (var originalInput = sample.OpenRead())
        {
            originalStreams = CompoundFile.ReadStreamsWithPaths(originalInput);
        }

        using var workbookInput = sample.OpenRead();
        using var workbook = new HSSFWorkbook(workbookInput);

        using var output = new MemoryStream();
        workbook.write(output);
        output.Position = 0;
        var writtenStreams = CompoundFile.ReadStreamsWithPaths(output);

        foreach (var streamName in originalStreams.Keys.Where(name => name != "Workbook"))
        {
            Assert.True(writtenStreams.ContainsKey(streamName), $"Missing preserved stream '{streamName}'.");
            Assert.Equal(originalStreams[streamName], writtenStreams[streamName]);
        }

        Assert.Contains("_VBA_PROJECT_CUR/VBA/dir", writtenStreams.Keys);
        Assert.Contains("_VBA_PROJECT_CUR/VBA/Module1", writtenStreams.Keys);
    }

    [Fact]
    public void Write_LoadedWorkbook_PreservesOleDirectoryEntryMetadata()
    {
        using var seedWorkbook = new HSSFWorkbook();
        seedWorkbook.createSheet("Data").createRow(0).createCell(0).setCellValue("alpha");
        using var seedStream = new MemoryStream();
        seedWorkbook.write(seedStream);
        seedStream.Position = 0;
        var seedStreams = CompoundFile.ReadStreamsWithPaths(seedStream);

        var metadata = new Dictionary<string, CompoundFileEntryMetadata>(StringComparer.Ordinal)
        {
            [""] = new(5, 1, Enumerable.Range(0, 16).Select(i => (byte)i).ToArray(), 7, 123456789, 987654321),
            ["Meta"] = new(1, 0, Enumerable.Range(16, 16).Select(i => (byte)i).ToArray(), 9, 223456789, 887654321),
            ["Meta/Leaf"] = new(2, 1, Enumerable.Range(32, 16).Select(i => (byte)i).ToArray(), 11, 323456789, 787654321)
        };
        seedStreams["Meta/Leaf"] = new byte[] { 1, 2, 3, 4 };

        using var source = new MemoryStream();
        CompoundFile.Write(source, new CompoundFileDocument(seedStreams, metadata));
        source.Position = 0;
        using var workbook = new HSSFWorkbook(source);

        using var output = new MemoryStream();
        workbook.write(output);
        output.Position = 0;
        var writtenDocument = CompoundFile.ReadDocument(output);

        AssertMetadataEqual(metadata[""], writtenDocument.EntryMetadata[""]);
        AssertMetadataEqual(metadata["Meta"], writtenDocument.EntryMetadata["Meta"]);
        AssertMetadataEqual(metadata["Meta/Leaf"], writtenDocument.EntryMetadata["Meta/Leaf"]);
    }

    [Fact]
    public void Write_LoadedWorkbook_PreservesUnknownBiffRecordsDuringLightEdit()
    {
        using var seedWorkbook = new HSSFWorkbook();
        seedWorkbook.createSheet("Data").createRow(0).createCell(0).setCellValue("alpha");
        using var seedStream = new MemoryStream();
        seedWorkbook.write(seedStream);
        seedStream.Position = 0;
        var seedDocument = CompoundFile.ReadDocument(seedStream);
        var workbookStream = seedDocument.Streams["Workbook"];
        var globalUnknown = CreateBiffRecord(0x0BAD, new byte[] { 1, 3, 5, 7 });
        var sheetUnknown = CreateBiffRecord(0x0BEE, new byte[] { 2, 4, 6, 8, 10 });
        seedDocument.Streams["Workbook"] = InjectUnknownBiffRecords(workbookStream, globalUnknown, sheetUnknown);

        using var source = new MemoryStream();
        CompoundFile.Write(source, seedDocument);
        source.Position = 0;
        using var workbook = new HSSFWorkbook(source);
        workbook.getSheetAt(0).getRow(0)!.getCell(0)!.setCellValue("beta");

        using var output = new MemoryStream();
        workbook.write(output);
        output.Position = 0;
        var writtenWorkbookStream = CompoundFile.ReadDocument(output).Streams["Workbook"];

        Assert.True(ContainsSequence(writtenWorkbookStream, globalUnknown), "Missing preserved global unknown BIFF record.");
        Assert.True(ContainsSequence(writtenWorkbookStream, sheetUnknown), "Missing preserved sheet unknown BIFF record.");

        output.Position = 0;
        using var reread = new HSSFWorkbook(output);
        Assert.Equal("beta", reread.getSheetAt(0).getRow(0)!.getCell(0)!.getStringCellValue());
    }

    [Fact]
    public void Write_BookUppercasePoiFixture_PreservesWorkbookStreamAlias()
    {
        var sample = FindRepoRoot().Combine("poi/test-data/spreadsheet/BOOK_in_capitals.xls");
        using var input = sample.OpenRead();
        using var workbook = new HSSFWorkbook(input);

        using var output = new MemoryStream();
        workbook.write(output);
        output.Position = 0;
        var writtenStreams = CompoundFile.ReadStreams(output);

        Assert.True(writtenStreams.ContainsKey("BOOK"));
        Assert.False(writtenStreams.ContainsKey("Workbook"));
        Assert.False(writtenStreams.ContainsKey("Book"));
        Assert.False(writtenStreams.ContainsKey("WORKBOOK"));
    }

    private static DirectoryInfo FindRepoRoot()
    {
        var dir = new DirectoryInfo(AppContext.BaseDirectory);
        while (dir is not null && !File.Exists(Path.Combine(dir.FullName, "DotnetPOI.sln")))
        {
            dir = dir.Parent;
        }

        return dir ?? throw new InvalidOperationException("Could not locate repository root.");
    }

    private static void AssertMetadataEqual(CompoundFileEntryMetadata expected, CompoundFileEntryMetadata actual)
    {
        Assert.Equal(expected.Type, actual.Type);
        Assert.Equal(expected.Color, actual.Color);
        Assert.Equal(expected.ClassId, actual.ClassId);
        Assert.Equal(expected.StateBits, actual.StateBits);
        Assert.Equal(expected.CreationTime, actual.CreationTime);
        Assert.Equal(expected.ModifiedTime, actual.ModifiedTime);
    }

    private static byte[] InjectUnknownBiffRecords(byte[] workbookStream, byte[] globalRecord, byte[] sheetRecord)
    {
        var records = ReadBiffRecords(workbookStream);
        var firstSheetOffset = records
            .Where(record => record.Sid == 0x0085)
            .Select(record => BinaryPrimitives.ReadUInt32LittleEndian(workbookStream.AsSpan(record.DataOffset)))
            .Min();
        var globalEof = records.First(record => record.Sid == 0x000A && record.Offset < firstSheetOffset);
        var sheetBof = records.First(record => record.Offset == firstSheetOffset && record.Sid == 0x0809);
        var sheetInsertOffset = sheetBof.Offset + sheetBof.TotalLength;
        var patched = workbookStream.ToArray();

        foreach (var boundSheet in records.Where(record => record.Sid == 0x0085))
        {
            var originalOffset = BinaryPrimitives.ReadUInt32LittleEndian(patched.AsSpan(boundSheet.DataOffset));
            BinaryPrimitives.WriteUInt32LittleEndian(patched.AsSpan(boundSheet.DataOffset), originalOffset + (uint)globalRecord.Length);
        }

        using var output = new MemoryStream();
        output.Write(patched, 0, globalEof.Offset);
        output.Write(globalRecord, 0, globalRecord.Length);
        output.Write(patched, globalEof.Offset, sheetInsertOffset - globalEof.Offset);
        output.Write(sheetRecord, 0, sheetRecord.Length);
        output.Write(patched, sheetInsertOffset, patched.Length - sheetInsertOffset);
        return output.ToArray();
    }

    private static byte[] CreateBiffRecord(ushort sid, byte[] data)
    {
        var record = new byte[4 + data.Length];
        BinaryPrimitives.WriteUInt16LittleEndian(record.AsSpan(0), sid);
        BinaryPrimitives.WriteUInt16LittleEndian(record.AsSpan(2), (ushort)data.Length);
        data.CopyTo(record.AsSpan(4));
        return record;
    }

    private static bool ContainsSequence(byte[] haystack, byte[] needle)
    {
        for (var i = 0; i <= haystack.Length - needle.Length; i++)
        {
            if (haystack.AsSpan(i, needle.Length).SequenceEqual(needle))
            {
                return true;
            }
        }

        return false;
    }

    private static List<TestBiffRecord> ReadBiffRecords(byte[] stream)
    {
        var records = new List<TestBiffRecord>();
        var offset = 0;
        while (offset + 4 <= stream.Length)
        {
            var sid = BinaryPrimitives.ReadUInt16LittleEndian(stream.AsSpan(offset));
            var length = BinaryPrimitives.ReadUInt16LittleEndian(stream.AsSpan(offset + 2));
            if (offset + 4 + length > stream.Length)
            {
                break;
            }

            records.Add(new TestBiffRecord(sid, offset, offset + 4, 4 + length));
            offset += 4 + length;
        }

        return records;
    }

    private sealed record TestBiffRecord(ushort Sid, int Offset, int DataOffset, int TotalLength);
}

internal static class DirectoryInfoExtensions
{
    public static FileInfo Combine(this DirectoryInfo directory, string relativePath) =>
        new(Path.Combine(directory.FullName, relativePath));
}
