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
    public void Write_CellStyles_RoundTrip()
    {
        // Phase 12 item 4: style round-trip — font/alignment/border/wrap/format
        using var workbook = new HSSFWorkbook();

        var boldFont = workbook.createFont();
        boldFont.setBold(true);
        boldFont.setFontName("Calibri");
        boldFont.setFontHeightInPoints(14);
        boldFont.setItalic(true);
        boldFont.setColor((short)0x000C); // red palette index

        var style1 = workbook.createCellStyle();
        style1.setFont(boldFont);
        style1.setAlignment(HorizontalAlignment.Center);
        style1.setWrapText(true);
        style1.setBorderBottom(BorderStyle.Thin);
        style1.setBorderTop(BorderStyle.Thin);
        style1.setDataFormat(workbook.createDataFormat().getFormat("0.00"));

        var style2 = workbook.createCellStyle();
        style2.setAlignment(HorizontalAlignment.Right);
        style2.setBorderLeft(BorderStyle.Medium);
        style2.setBorderRight(BorderStyle.Thin);
        style2.setFillForegroundColor((short)3);
        style2.setFillPattern(FillPatternType.SolidForeground);

        var sheet = workbook.createSheet("Styles");
        var row0 = sheet.createRow(0);
        var cell0 = row0.createCell(0);
        cell0.setCellValue(42.5);
        cell0.setCellStyle(style1);

        var cell1 = row0.createCell(1);
        cell1.setCellValue("right aligned");
        cell1.setCellStyle(style2);

        // Cell with default style
        row0.createCell(2).setCellValue("default");

        using var stream = new MemoryStream();
        workbook.write(stream);
        stream.Position = 0;

        using var read = new HSSFWorkbook(stream);
        var rSheet = read.getSheetAt(0);
        var rRow = rSheet.getRow(0)!;

        // Verify cell0 style
        var rCell0 = rRow.getCell(0)!;
        Assert.Equal(CellType.Numeric, rCell0.getCellType());
        var rStyle1 = (HSSFCellStyle)rCell0.getCellStyle();
        Assert.Equal(HorizontalAlignment.Center, rStyle1.getAlignment());
        Assert.True(rStyle1.getWrapText());
        Assert.Equal(BorderStyle.Thin, rStyle1.getBorderBottom());
        Assert.Equal(BorderStyle.Thin, rStyle1.getBorderTop());
        var rFont1 = (HSSFFont)rStyle1.getFont();
        Assert.True(rFont1.getBold());
        Assert.True(rFont1.getItalic());
        Assert.Equal("Calibri", rFont1.getFontName());
        Assert.Equal(14, rFont1.getFontHeightInPoints());

        // Verify cell1 style
        var rCell1 = rRow.getCell(1)!;
        var rStyle2 = (HSSFCellStyle)rCell1.getCellStyle();
        Assert.Equal(HorizontalAlignment.Right, rStyle2.getAlignment());
        Assert.Equal(BorderStyle.Medium, rStyle2.getBorderLeft());
        Assert.Equal(BorderStyle.Thin, rStyle2.getBorderRight());

        // Verify cell2 default style
        var rCell2 = rRow.getCell(2)!;
        var rStyleDef = (HSSFCellStyle)rCell2.getCellStyle();
        Assert.Equal(HorizontalAlignment.General, rStyleDef.getAlignment());
        Assert.False(rStyleDef.getWrapText());
    }

    [Fact]
    public void Write_UserDefinedNumberFormat_RoundTrip()
    {
        // Phase 12 item 4: FormatRecord (0x041E) round-trip
        using var workbook = new HSSFWorkbook();
        var df = workbook.createDataFormat();
        var fmtIdx1 = df.getFormat("0.000");
        var fmtIdx2 = df.getFormat("yyyy-mm-dd");
        var fmtIdx3 = df.getFormat("#,##0.00 \"€\"");

        var style = workbook.createCellStyle();
        style.setDataFormat(fmtIdx2);

        var sheet = workbook.createSheet("Formats");
        var row = sheet.createRow(0);
        var cell = row.createCell(0);
        cell.setCellValue(44927.0); // numeric date
        cell.setCellStyle(style);

        using var ms = new MemoryStream();
        workbook.write(ms);
        ms.Position = 0;

        using var read = new HSSFWorkbook(ms);
        var rDf = read.createDataFormat();

        // Custom formats must survive round-trip
        Assert.Equal("0.000", rDf.getFormat(fmtIdx1));
        Assert.Equal("yyyy-mm-dd", rDf.getFormat(fmtIdx2));
        Assert.Equal("#,##0.00 \"€\"", rDf.getFormat(fmtIdx3));

        // Cell style should reference the correct format
        var rCell = read.getSheetAt(0).getRow(0)!.getCell(0)!;
        var rStyle = (HSSFCellStyle)rCell.getCellStyle();
        Assert.Equal(fmtIdx2, rStyle.getDataFormat());
        Assert.Equal("yyyy-mm-dd", rStyle.getDataFormatString());
    }

    [Fact]
    public void Write_AllCellTypes_RoundTrip()
    {
        // Phase 12 item 3: C# round-trip for all 5 cell types
        using var workbook = new HSSFWorkbook();
        var sheet = workbook.createSheet("Types");
        var row = sheet.createRow(0);

        row.createCell(0).setCellValue("hello world");   // String
        row.createCell(1).setCellValue(3.14159);         // Numeric
        row.createCell(2).setCellValue(true);            // Boolean true
        row.createCell(3).setCellValue(false);           // Boolean false
        row.createCell(4).setCellErrorValue(0x07);       // Error: #DIV/0!
        row.createCell(5).setCellErrorValue(0x2A);       // Error: #N/A
        row.createCell(6);                               // Blank (no value set)

        using var stream = new MemoryStream();
        workbook.write(stream);
        stream.Position = 0;

        using var read = new HSSFWorkbook(stream);
        var readRow = read.getSheetAt(0).getRow(0);
        Assert.NotNull(readRow);

        // String
        var c0 = readRow!.getCell(0)!;
        Assert.Equal(CellType.String, c0.getCellType());
        Assert.Equal("hello world", c0.getStringCellValue());

        // Numeric
        var c1 = readRow.getCell(1)!;
        Assert.Equal(CellType.Numeric, c1.getCellType());
        Assert.Equal(3.14159, c1.getNumericCellValue(), 5);

        // Boolean true
        var c2 = readRow.getCell(2)!;
        Assert.Equal(CellType.Boolean, c2.getCellType());
        Assert.True(c2.getBooleanCellValue());

        // Boolean false
        var c3 = readRow.getCell(3)!;
        Assert.Equal(CellType.Boolean, c3.getCellType());
        Assert.False(c3.getBooleanCellValue());

        // Error #DIV/0!
        var c4 = readRow.getCell(4)!;
        Assert.Equal(CellType.Error, c4.getCellType());
        Assert.Equal("#DIV/0!", c4.getErrorCellString());

        // Error #N/A
        var c5 = readRow.getCell(5)!;
        Assert.Equal(CellType.Error, c5.getCellType());
        Assert.Equal("#N/A", c5.getErrorCellString());

        // Blank
        var c6 = readRow.getCell(6)!;
        Assert.Equal(CellType.Blank, c6.getCellType());
    }

    [Fact]
    public void Write_MultipleSheets_RoundTrip()
    {
        // Phase 12 item 3: multiple sheets, sparse rows, sparse cells, high column index
        using var workbook = new HSSFWorkbook();

        // Sheet 1: sparse rows and columns
        var sheet1 = workbook.createSheet("First");
        sheet1.createRow(0).createCell(0).setCellValue("A1");
        sheet1.createRow(2).createCell(3).setCellValue("sparse"); // non-contiguous row/col
        sheet1.createRow(10).createCell(0).setCellValue(99.9);

        // Sheet 2: empty rows (row exists, no cells)
        var sheet2 = workbook.createSheet("Second");
        sheet2.createRow(0); // empty row - row exists but no cells
        sheet2.createRow(1).createCell(0).setCellValue(true);

        // Sheet 3: high column index (near BIFF8 max of 255)
        var sheet3 = workbook.createSheet("Third");
        sheet3.createRow(0).createCell(255).setCellValue("last col"); // column index 255 (0-based)
        sheet3.createRow(0).createCell(0).setCellValue("first col");

        using var stream = new MemoryStream();
        workbook.write(stream);
        stream.Position = 0;

        using var read = new HSSFWorkbook(stream);
        Assert.Equal(3, read.getNumberOfSheets());

        // Verify sheet 1 by name lookup and cell values
        var r1 = read.getSheet("First");
        Assert.NotNull(r1);
        Assert.Equal("A1", r1!.getRow(0)!.getCell(0)!.getStringCellValue());
        Assert.Equal("sparse", r1.getRow(2)!.getCell(3)!.getStringCellValue());
        Assert.Equal(99.9, r1.getRow(10)!.getCell(0)!.getNumericCellValue());

        // Verify sheet 2
        var r2 = read.getSheet("Second");
        Assert.NotNull(r2);
        Assert.True(r2!.getRow(1)!.getCell(0)!.getBooleanCellValue());

        // Verify sheet 3 (high column index)
        var r3 = read.getSheet("Third");
        Assert.NotNull(r3);
        Assert.Equal("first col", r3!.getRow(0)!.getCell(0)!.getStringCellValue());
        Assert.Equal("last col", r3.getRow(0)!.getCell(255)!.getStringCellValue());
    }

    [Fact]
    public void Write_SheetLayout_RoundTrip()
    {
        // Phase 12 item 4: layout round-trip — column width, row height, merge, freeze, hidden
        using var workbook = new HSSFWorkbook();
        var sheet = workbook.createSheet("Layout");

        sheet.setColumnWidth(0, 5000);
        sheet.setColumnWidth(1, 8000);
        sheet.setColumnHidden(2, true);

        var row0 = sheet.createRow(0);
        row0.setHeight(30.0f);
        row0.createCell(0).setCellValue("tall row");

        var row1 = sheet.createRow(1);
        row1.setHidden(true);
        row1.createCell(0).setCellValue("hidden");

        sheet.addMergedRegion(new SS.Util.CellRangeAddress(2, 2, 0, 2));
        sheet.createRow(2).createCell(0).setCellValue("merged");

        sheet.createFreezePane(1, 2);

        using var ms = new MemoryStream();
        workbook.write(ms);
        ms.Position = 0;

        using var read = new HSSFWorkbook(ms);
        var rSheet = read.getSheetAt(0);

        // Column widths
        Assert.Equal(5000, rSheet.getColumnWidth(0));
        Assert.Equal(8000, rSheet.getColumnWidth(1));
        Assert.True(rSheet.isColumnHidden(2));

        // Row height
        var rRow0 = rSheet.getRow(0)!;
        Assert.True(rRow0.getHeight() > 20f, "Row 0 height should be > 20pt.");

        // Hidden row (hidden via setZeroHeight — row height = 0 means hidden in BIFF)
        var rRow1 = rSheet.getRow(1)!;
        Assert.True(rRow1.isHidden());

        // Merged region
        var regions = rSheet.getMergedRegions();
        Assert.Single(regions);
        Assert.Equal(2, regions[0].FirstRow);
        Assert.Equal(2, regions[0].LastRow);
        Assert.Equal(0, regions[0].FirstCol);
        Assert.Equal(2, regions[0].LastCol);

        // Freeze pane
        Assert.Equal(1, rSheet.FreezeColSplit);
        Assert.Equal(2, rSheet.FreezeRowSplit);
    }

    [Fact]
    public void Write_RecordLevelIntegrity_SstLabelSstNumberBoolErrBlankDimensions()
    {
        // Phase 12 item 3: verify SST/LabelSST/Number/BoolErr/Blank/Dimensions/BoundSheet record integrity
        using var workbook = new HSSFWorkbook();
        var sheet = workbook.createSheet("Sheet1");
        var row = sheet.createRow(0);
        row.createCell(0).setCellValue("str1");   // → LabelSST
        row.createCell(1).setCellValue("str2");   // → LabelSST (different SST index)
        row.createCell(2).setCellValue("str1");   // → LabelSST (same SST index as col 0)
        row.createCell(3).setCellValue(1.5);      // → Number
        row.createCell(4).setCellValue(true);     // → BoolErr (fError=0)
        row.createCell(4).setCellErrorValue(0x07); // overwrite with error → BoolErr (fError=1)
        row.createCell(5);                        // → Blank

        using var ms = new MemoryStream();
        workbook.write(ms);
        var bytes = ms.ToArray();

        // Parse the OLE2 compound file and extract the Workbook stream
        using var cfStream = new MemoryStream(bytes);
        var streams = CompoundFile.ReadStreams(cfStream);
        Assert.True(streams.ContainsKey("Workbook"), "Missing 'Workbook' stream.");
        var wbStream = streams["Workbook"];

        var records = ReadBiffRecords(wbStream);
        Assert.True(records.Count > 0, "No BIFF records parsed.");

        // Verify BOF record (0x0809) is the first record
        Assert.Equal((ushort)0x0809, records[0].Sid);

        // Verify SST record exists and has the correct unique count
        var sstRecords = records.Where(r => r.Sid == 0x00FC).ToList();
        Assert.Single(sstRecords);
        var sstData = sstRecords[0].Data;
        Assert.True(sstData.Length >= 8, "SST record too short.");
        var uniqueCount = System.Buffers.Binary.BinaryPrimitives.ReadUInt32LittleEndian(sstData.AsSpan(4));
        Assert.Equal(2u, uniqueCount); // "str1" and "str2" are unique

        // Verify BoundSheet8 record (0x0085) exists with correct offset
        var boundSheets = records.Where(r => r.Sid == 0x0085).ToList();
        Assert.Single(boundSheets);
        var bsData = boundSheets[0].Data;
        Assert.True(bsData.Length >= 8, "BoundSheet record too short.");
        var sheetOffset = System.Buffers.Binary.BinaryPrimitives.ReadUInt32LittleEndian(bsData.AsSpan(0));
        Assert.True(sheetOffset > 0, "BoundSheet sheet offset must be non-zero.");

        // Verify the BoundSheet offset points to a BOF record (0x0809)
        Assert.True(sheetOffset + 4 <= wbStream.Length, "BoundSheet offset is beyond end of stream.");
        var recordAtOffset = System.Buffers.Binary.BinaryPrimitives.ReadUInt16LittleEndian(wbStream.AsSpan((int)sheetOffset));
        Assert.Equal((ushort)0x0809, recordAtOffset);

        // Verify Dimensions record (0x0200) exists in the sheet
        var dims = records.Where(r => r.Sid == 0x0200).ToList();
        Assert.NotEmpty(dims);
        var dimData = dims[0].Data;
        Assert.True(dimData.Length >= 14, "Dimensions record too short.");

        // Verify LabelSST records (0x00FD) reference valid SST indices
        var labelSstRecords = records.Where(r => r.Sid == 0x00FD).ToList();
        Assert.Equal(3, labelSstRecords.Count); // col 0, 1, 2 are strings
        foreach (var lsst in labelSstRecords)
        {
            Assert.True(lsst.Data.Length >= 10, "LabelSST record too short.");
            var sstIdx = System.Buffers.Binary.BinaryPrimitives.ReadUInt32LittleEndian(lsst.Data.AsSpan(6));
            Assert.True(sstIdx < uniqueCount, $"LabelSST SST index {sstIdx} exceeds unique count {uniqueCount}.");
        }

        // Verify Number record (0x0203) for col 3
        var numRecords = records.Where(r => r.Sid == 0x0203).ToList();
        Assert.Single(numRecords);
        Assert.True(numRecords[0].Data.Length >= 14, "Number record too short.");

        // Verify BoolErr records (0x0205) for col 4 (error)
        var boolErrRecords = records.Where(r => r.Sid == 0x0205).ToList();
        Assert.Single(boolErrRecords);
        Assert.True(boolErrRecords[0].Data.Length >= 8, "BoolErr record too short.");

        // Verify Blank record (0x0201) for col 5
        var blankRecords = records.Where(r => r.Sid == 0x0201).ToList();
        Assert.Single(blankRecords);
        Assert.True(blankRecords[0].Data.Length >= 6, "Blank record too short.");
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
    public void LightEdit_PoiSimpleMultiCellFixture_UpdatesCellAndPreservesUnknownRecords()
    {
        // Phase 12 item 3: light-edit an existing POI fixture and verify cell update + unknown BIFF preservation
        var sample = FindRepoRoot().Combine("poi/test-data/spreadsheet/SimpleMultiCell.xls");
        using var originalInput = sample.OpenRead();
        using var workbook = new HSSFWorkbook(originalInput);

        // Edit a cell in the first sheet
        var sheet = workbook.getSheetAt(0);
        var targetRow = sheet.getRow(0) ?? sheet.createRow(0);
        var targetCell = targetRow.getCell(0) ?? targetRow.createCell(0);
        targetCell.setCellValue("EDITED_BY_DOTNET_POI");

        using var output = new MemoryStream();
        workbook.write(output);
        output.Position = 0;

        // Verify the edit round-tripped
        using var reread = new HSSFWorkbook(output);
        var readSheet = reread.getSheetAt(0);
        var readCell = readSheet.getRow(0)?.getCell(0);
        Assert.NotNull(readCell);
        Assert.Equal("EDITED_BY_DOTNET_POI", readCell!.getStringCellValue());

        // Verify the workbook still has the same number of sheets
        Assert.Equal(workbook.getNumberOfSheets(), reread.getNumberOfSheets());

        // Verify the output is a valid OLE2 compound file with a Workbook stream
        output.Position = 0;
        var writtenStreams = CompoundFile.ReadStreams(output);
        Assert.True(writtenStreams.ContainsKey("Workbook") || writtenStreams.ContainsKey("WORKBOOK") || writtenStreams.ContainsKey("Book"),
            "Written file must have a Workbook/Book stream.");
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

            var data = stream.AsSpan(offset + 4, length).ToArray();
            records.Add(new TestBiffRecord(sid, offset, offset + 4, 4 + length, data));
            offset += 4 + length;
        }

        return records;
    }

    private sealed record TestBiffRecord(ushort Sid, int Offset, int DataOffset, int TotalLength, byte[] Data);
}

internal static class DirectoryInfoExtensions
{
    public static FileInfo Combine(this DirectoryInfo directory, string relativePath) =>
        new(Path.Combine(directory.FullName, relativePath));
}
