using System.Globalization;
using System.IO.Compression;
using System.Text;
using System.Xml;
using DotnetPoi.POIFS.Crypt;
using DotnetPoi.SS.UserModel;
using DotnetPoi.SS.Util;
using DotnetPoi.SS.Xml;

namespace DotnetPoi.XSSF.UserModel;

public sealed class XSSFWorkbook : IWorkbook
{
    public const int PICTURE_TYPE_EMF = 2;
    public const int PICTURE_TYPE_WMF = 3;
    public const int PICTURE_TYPE_PICT = 4;
    public const int PICTURE_TYPE_JPEG = 5;
    public const int PICTURE_TYPE_PNG = 6;
    public const int PICTURE_TYPE_DIB = 7;
    public const int PICTURE_TYPE_GIF = 8;
    public const int PICTURE_TYPE_TIFF = 9;
    public const int PICTURE_TYPE_EPS = 10;
    public const int PICTURE_TYPE_BMP = 11;
    public const int PICTURE_TYPE_WPG = 12;

    private readonly List<XSSFSheet> _sheets = new();
    private readonly List<XSSFPictureData> _pictures = new();
    private readonly Dictionary<string, int> _sharedStringIndexes = new(StringComparer.Ordinal);
    private readonly List<XSSFRichTextString> _sharedStrings = new();
    private readonly List<XSSFFont> _fonts = new();
    private readonly List<XSSFCellStyle> _cellStyles = new();
    private readonly List<XSSFCellStyle?> _fills = new();
    private readonly List<XSSFCellStyle?> _borders = new();
    private readonly Dictionary<string, int> _customNumberFormatIndexes = new(StringComparer.Ordinal);
    private readonly SortedDictionary<int, string> _customNumberFormats = new();
    private XSSFCreationHelper? _creationHelper;
    private XSSFDataFormat? _dataFormat;
    private int _nextDrawingIndex = 1;
    private bool _hasCalcPr;
    private bool _forceFormulaRecalculation;

    // xlsm support: opaque VBA binary preserved byte-for-byte.
    // Non-null iff the loaded/constructed workbook is macro-enabled.
    private byte[]? _vbaProjectBin;

    // Pivot table support
    private readonly List<XSSFPivotTable> _pivotTables = new();
    private int _nextPivotCacheId;

    // Unknown part preservation: ZIP entries not understood by the model
    // (pivot tables, charts, etc.) are stored byte-for-byte during Load()
    // and re-emitted during write() before model parts overwrite.
    private readonly Dictionary<string, byte[]> _preservedEntries = new(StringComparer.OrdinalIgnoreCase);

    // Workbook protection
    private bool _workbookProtected;

    private const string ContentTypeXlsx = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml";
    private const string ContentTypeXlsm = "application/vnd.ms-excel.sheet.macroEnabled.main+xml";
    private const string ContentTypeVbaProject = "application/vnd.ms-office.vbaProject";
    private const string RelTypeVbaProject = "http://schemas.microsoft.com/office/2006/relationships/vbaProject";

    /// <summary>True if this workbook was loaded from xlsm (has a vbaProject.bin).</summary>
    public bool HasMacros => _vbaProjectBin != null;

    public bool isMacroEnabled() => HasMacros;

    public XSSFWorkbook()
    {
        InitializeDefaultStyles();
    }

    public XSSFWorkbook(Stream stream)
    {
        ArgumentNullException.ThrowIfNull(stream);
        InitializeDefaultStyles();
        Load(stream);
    }

    public XSSFSheet createSheet()
    {
        return createSheet("Sheet" + (_sheets.Count + 1).ToString(CultureInfo.InvariantCulture));
    }

    public XSSFSheet createSheet(string sheetname)
    {
        if (string.IsNullOrEmpty(sheetname))
        {
            throw new ArgumentException("Sheet name must not be empty.", nameof(sheetname));
        }

        var sheet = new XSSFSheet(this, sheetname, _sheets.Count + 1);
        _sheets.Add(sheet);
        return sheet;
    }

    public XSSFSheet getSheetAt(int index)
    {
        return _sheets[index];
    }

    public XSSFSheet? getSheet(string name)
    {
        return _sheets.FirstOrDefault(sheet => string.Equals(sheet.SheetName, name, StringComparison.Ordinal));
    }

    public int getNumberOfSheets()
    {
        return _sheets.Count;
    }

    internal IReadOnlyList<XSSFSheet> Sheets => _sheets;

    public XSSFCreationHelper getCreationHelper()
    {
        return _creationHelper ??= new XSSFCreationHelper(this);
    }

    public XSSFCellStyle createCellStyle()
    {
        var style = new XSSFCellStyle(this, _cellStyles.Count);
        _cellStyles.Add(style);
        return style;
    }

    public XSSFCellStyle getCellStyleAt(int idx)
    {
        return _cellStyles[idx];
    }

    public XSSFDataFormat createDataFormat()
    {
        return _dataFormat ??= new XSSFDataFormat(this);
    }

    public XSSFFont createFont()
    {
        var font = new XSSFFont(_fonts.Count);
        _fonts.Add(font);
        return font;
    }

    public XSSFFont getFontAt(int idx)
    {
        return _fonts[idx];
    }

    public int addPicture(byte[] pictureData, int format)
    {
        ArgumentNullException.ThrowIfNull(pictureData);
        ValidatePictureType(format);

        var picture = new XSSFPictureData(pictureData, format, _pictures.Count + 1);
        _pictures.Add(picture);
        return picture.Index - 1;
    }

    public int addPicture(Stream stream, int format)
    {
        ArgumentNullException.ThrowIfNull(stream);
        using var memory = new MemoryStream();
        stream.CopyTo(memory);
        return addPicture(memory.ToArray(), format);
    }

    public IReadOnlyList<XSSFPictureData> getAllPictures()
    {
        return _pictures;
    }

    /// <summary>
    /// Tells Excel to recalculate all formulas when the workbook is opened.
    /// Ported from XSSFWorkbook.setForceFormulaRecalculation(boolean).
    /// </summary>
    public void setForceFormulaRecalculation(bool value)
    {
        _hasCalcPr = true;
        _forceFormulaRecalculation = value;
    }

    /// <summary>
    /// Returns whether Excel will be asked to recalculate all formulas on open.
    /// Ported from XSSFWorkbook.getForceFormulaRecalculation().
    /// </summary>
    public bool getForceFormulaRecalculation() => _hasCalcPr && _forceFormulaRecalculation;

    IReadOnlyList<IPictureData> IWorkbook.getAllPictures() => _pictures;

    ISheet IWorkbook.createSheet() => createSheet();

    ISheet IWorkbook.createSheet(string sheetname) => createSheet(sheetname);

    ISheet IWorkbook.getSheetAt(int index) => getSheetAt(index);

    ISheet? IWorkbook.getSheet(string name) => getSheet(name);

    ICreationHelper IWorkbook.getCreationHelper() => getCreationHelper();

    ICellStyle IWorkbook.createCellStyle() => createCellStyle();

    ICellStyle IWorkbook.getCellStyleAt(int idx) => getCellStyleAt(idx);

    IDataFormat IWorkbook.createDataFormat() => createDataFormat();

    IFont IWorkbook.createFont() => createFont();

    IFont IWorkbook.getFontAt(int idx) => getFontAt(idx);

    void IWorkbook.setForceFormulaRecalculation(bool value) => setForceFormulaRecalculation(value);

    bool IWorkbook.getForceFormulaRecalculation() => getForceFormulaRecalculation();

    public void protectWorkbook(bool protect) => _workbookProtected = protect;

    public bool isWorkbookProtected() => _workbookProtected;

    void IWorkbook.protectWorkbook(bool protect) => protectWorkbook(protect);

    bool IWorkbook.isWorkbookProtected() => isWorkbookProtected();

    public void write(Stream stream)
    {
        ArgumentNullException.ThrowIfNull(stream);
        EnsureAtLeastOneSheet();
        BuildSharedStrings();
        AssignHyperlinkRelationshipIds();

        using var archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true);

        // Emit preserved (unknown) entries first so model entries can overwrite
        foreach (var kv in _preservedEntries)
        {
            WriteBinaryEntry(archive, kv.Key, kv.Value);
        }

        WriteEntry(archive, "[Content_Types].xml", WriteContentTypes);
        WriteEntry(archive, "_rels/.rels", WriteRootRelationships);
        WriteEntry(archive, "docProps/app.xml", WriteAppProperties);
        WriteEntry(archive, "docProps/core.xml", WriteCoreProperties);
        WriteEntry(archive, "xl/workbook.xml", WriteWorkbook);
        WriteEntry(archive, "xl/_rels/workbook.xml.rels", WriteWorkbookRelationships);
        WriteEntry(archive, "xl/styles.xml", WriteStyles);
        WriteEntry(archive, "xl/sharedStrings.xml", WriteSharedStrings);

        foreach (var sheet in _sheets)
        {
            WriteEntry(archive, $"xl/worksheets/sheet{sheet.SheetIndex}.xml", writer => WriteWorksheet(writer, sheet));
            bool hasRels = sheet.Drawing is not null || sheet.Hyperlinks.Count > 0 || sheet.PivotTables.Count > 0;
            if (hasRels)
            {
                WriteEntry(archive, $"xl/worksheets/_rels/sheet{sheet.SheetIndex}.xml.rels", writer => WriteSheetRelationships(writer, sheet));
            }
            if (sheet.Drawing is not null)
            {
                WriteEntry(archive, $"xl/drawings/drawing{sheet.Drawing.DrawingIndex}.xml", writer => WriteDrawing(writer, sheet.Drawing));
                WriteEntry(archive, $"xl/drawings/_rels/drawing{sheet.Drawing.DrawingIndex}.xml.rels", writer => WriteDrawingRelationships(writer, sheet.Drawing));
            }
        }

        foreach (var picture in _pictures)
        {
            WriteBinaryEntry(archive, $"xl/media/image{picture.Index}.{picture.Extension}", picture.Data);
        }

        if (_vbaProjectBin != null)
            WriteBinaryEntry(archive, "xl/vbaProject.bin", _vbaProjectBin);

        // Pivot table parts
        foreach (var pt in _pivotTables)
        {
            var ptIndex = pt.PivotTableIndex;
            var cacheId = pt.CacheId;
            WriteEntry(archive, $"xl/pivotTables/pivotTable{ptIndex}.xml", pt.WritePivotTableDefinition);
            WriteEntry(archive, $"xl/pivotCache/pivotCacheDefinition{cacheId + 1}.xml", writer => WritePivotCacheDefinition(writer, pt.CacheDefinition));
            WriteEntry(archive, $"xl/pivotCache/pivotCacheRecords{cacheId + 1}.xml", writer => WritePivotCacheRecords(writer, pt.CacheRecords));
        }
    }

    public void writeEncrypted(Stream stream, string password)
    {
        ArgumentNullException.ThrowIfNull(stream);
        using var package = new MemoryStream();
        write(package);

        var info = new EncryptionInfo(EncryptionMode.agile);
        info.Encryptor.confirmPassword(password);
        info.Encryptor.encryptPackage(package.ToArray(), stream);
    }

    public void close()
    {
    }

    public void setVBAProject(byte[] vbaProjectData)
    {
        ArgumentNullException.ThrowIfNull(vbaProjectData);
        _vbaProjectBin = vbaProjectData.ToArray();
    }

    public void setVBAProject(Stream vbaProjectStream)
    {
        ArgumentNullException.ThrowIfNull(vbaProjectStream);
        using var ms = new MemoryStream();
        vbaProjectStream.CopyTo(ms);
        _vbaProjectBin = ms.ToArray();
    }

    public void Dispose()
    {
        close();
    }

    private void EnsureAtLeastOneSheet()
    {
        if (_sheets.Count == 0)
        {
            createSheet();
        }
    }

    private void InitializeDefaultStyles()
    {
        _fonts.Clear();
        _cellStyles.Clear();
        _fills.Clear();
        _borders.Clear();
        _customNumberFormatIndexes.Clear();
        _customNumberFormats.Clear();
        _dataFormat = null;

        _fonts.Add(new XSSFFont(0));
        _fills.Add(null);
        _fills.Add(null);
        _borders.Add(null);
        _cellStyles.Add(new XSSFCellStyle(this, 0));
    }

    internal void VerifyFontBelongsToWorkbook(XSSFFont font)
    {
        if ((uint)font.getIndex() >= (uint)_fonts.Count || !ReferenceEquals(_fonts[font.getIndex()], font))
        {
            throw new ArgumentException("This Font does not belong to this Workbook.");
        }
    }

    internal int GetOrAddCustomNumberFormat(string format)
    {
        if (_customNumberFormatIndexes.TryGetValue(format, out var existingIndex))
        {
            return existingIndex;
        }

        var nextIndex = _customNumberFormats.Count == 0
            ? XSSFDataFormat.FIRST_USER_DEFINED_FORMAT_INDEX
            : Math.Max(XSSFDataFormat.FIRST_USER_DEFINED_FORMAT_INDEX, _customNumberFormats.Keys.Max() + 1);
        _customNumberFormats[nextIndex] = format;
        _customNumberFormatIndexes[format] = nextIndex;
        return nextIndex;
    }

    internal string? GetCustomNumberFormat(int index)
    {
        return _customNumberFormats.TryGetValue(index, out var format) ? format : null;
    }

    internal int GetOrAddFill(XSSFCellStyle style)
    {
        _fills.Add(style);
        return _fills.Count - 1;
    }

    internal int GetOrAddBorder(XSSFCellStyle style)
    {
        _borders.Add(style);
        return _borders.Count - 1;
    }

    internal void AssignHyperlinkRelationshipIds()
    {
        int nextId = 1;
        foreach (var sheet in _sheets)
        {
            if (sheet.Drawing is not null) nextId = 2;
            foreach (var hyperlink in sheet.Hyperlinks)
            {
                hyperlink.RelationshipId = $"rId{nextId++}";
            }
        }
    }

    internal int GetNextDrawingIndex()
    {
        return _nextDrawingIndex++;
    }

    internal int AllocatePivotCacheId()
    {
        return _nextPivotCacheId++;
    }

    internal void RegisterPivotTable(XSSFPivotTable pivotTable)
    {
        _pivotTables.Add(pivotTable);
    }

    internal IReadOnlyList<XSSFPivotTable> PivotTables => _pivotTables;

    internal XSSFPictureData GetPictureData(int pictureIndex)
    {
        if ((uint)pictureIndex >= (uint)_pictures.Count)
        {
            throw new ArgumentException("Picture index does not exist in this workbook.", nameof(pictureIndex));
        }

        return _pictures[pictureIndex];
    }

    private void BuildSharedStrings()
    {
        _sharedStringIndexes.Clear();
        _sharedStrings.Clear();

        foreach (var sheet in _sheets)
        {
            foreach (var row in sheet.Rows)
            {
                foreach (var cell in row.Cells)
                {
                    if (cell.getCellType() != CellType.String) continue;

                    var rts = cell.getRichStringCellValue();
                    var plainText = rts.getString();
                    if (_sharedStringIndexes.ContainsKey(plainText)) continue;

                    _sharedStringIndexes[plainText] = _sharedStrings.Count;
                    _sharedStrings.Add(rts);
                }
            }
        }
    }

    private int GetSharedStringIndex(string value)
    {
        return _sharedStringIndexes[value];
    }

    private static void WriteEntry(ZipArchive archive, string name, Action<PoiXmlWriter> write)
    {
        var entry = archive.CreateEntry(name, CompressionLevel.Optimal);
        using var entryStream = entry.Open();
        using var textWriter = new StreamWriter(entryStream, new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
        using var writer = PoiXmlWriterFactory.CreateForOoxmlPackagePart(textWriter, name);
        write(writer);
    }

    private static void WriteBinaryEntry(ZipArchive archive, string name, byte[] data)
    {
        var entry = archive.CreateEntry(name, CompressionLevel.Optimal);
        using var entryStream = entry.Open();
        entryStream.Write(data, 0, data.Length);
    }

    private void Load(Stream stream)
    {
        _sheets.Clear();
        _vbaProjectBin = null;
        _hasCalcPr = false;
        _forceFormulaRecalculation = false;
        InitializeDefaultStyles();

        using var archive = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: true);
        var sharedStrings = ReadSharedStrings(archive);
        ReadStyles(archive);
        ReadPictures(archive);
        ReadVbaProject(archive);
        ReadWorkbookCalcPr(archive);

        foreach (var sheetInfo in ReadWorkbookSheets(archive))
        {
            var sheet = createSheet(sheetInfo.Name);
            ReadWorksheet(archive, sheetInfo.PartName, sheet, sharedStrings);
            ReadSheetDrawing(archive, sheetInfo.PartName, sheet);
        }

        CollectPreservedEntries(archive);
    }

    private void ReadWorkbookCalcPr(ZipArchive archive)
    {
        var workbookEntry = archive.GetEntry("xl/workbook.xml")
            ?? throw new InvalidDataException("The xlsx package is missing xl/workbook.xml.");

        using var workbookStream = workbookEntry.Open();
        using var reader = XmlReader.Create(workbookStream, new XmlReaderSettings { IgnoreWhitespace = false });
        while (reader.Read())
        {
            if (reader.NodeType != XmlNodeType.Element)
                continue;

            if (reader.LocalName == "calcPr")
            {
                _hasCalcPr = true;
                _forceFormulaRecalculation = ParseBooleanAttribute(reader.GetAttribute("fullCalcOnLoad"));
            }
            else if (reader.LocalName == "workbookProtection")
            {
                _workbookProtected = true;
            }
        }
    }

    private static bool ParseBooleanAttribute(string? value) =>
        string.Equals(value, "true", StringComparison.OrdinalIgnoreCase)
        || string.Equals(value, "1", StringComparison.Ordinal);

    private void ReadVbaProject(ZipArchive archive)
    {
        var entry = archive.GetEntry("xl/vbaProject.bin");
        if (entry is null) return;
        using var s = entry.Open();
        using var ms = new MemoryStream();
        s.CopyTo(ms);
        _vbaProjectBin = ms.ToArray();
    }

    /// <summary>
    /// Collects ZIP entries not understood by the model and stores them
    /// byte-for-byte so they survive round-trip (pivot tables, charts, etc.).
    /// Ported from XWPFDocument / XMLSlideShow approach.
    /// </summary>
    private void CollectPreservedEntries(ZipArchive archive)
    {
        _preservedEntries.Clear();
        var known = GetModelEntryNames();
        foreach (var entry in archive.Entries)
        {
            var name = entry.FullName.Replace('\\', '/');
            if (known.Contains(name, StringComparer.OrdinalIgnoreCase)) continue;
            if (name.StartsWith("xl/media/", StringComparison.OrdinalIgnoreCase)) continue;
            using var s = entry.Open();
            using var ms = new MemoryStream();
            s.CopyTo(ms);
            _preservedEntries[name] = ms.ToArray();
        }
    }

    /// <summary>Returns the set of entry paths that the model knows how to write.</summary>
    private HashSet<string> GetModelEntryNames()
    {
        var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "[Content_Types].xml",
            "_rels/.rels",
            "docProps/app.xml",
            "docProps/core.xml",
            "xl/workbook.xml",
            "xl/_rels/workbook.xml.rels",
            "xl/styles.xml",
            "xl/sharedStrings.xml",
        };

        foreach (var sheet in _sheets)
        {
            names.Add($"xl/worksheets/sheet{sheet.SheetIndex}.xml");
            bool hasRels = sheet.Drawing is not null || sheet.Hyperlinks.Count > 0;
            if (hasRels)
                names.Add($"xl/worksheets/_rels/sheet{sheet.SheetIndex}.xml.rels");
            if (sheet.Drawing is not null)
            {
                names.Add($"xl/drawings/drawing{sheet.Drawing.DrawingIndex}.xml");
                names.Add($"xl/drawings/_rels/drawing{sheet.Drawing.DrawingIndex}.xml.rels");
            }
        }

        foreach (var picture in _pictures)
        {
            names.Add($"xl/media/image{picture.Index}.{picture.Extension}");
        }

        if (_vbaProjectBin != null)
            names.Add("xl/vbaProject.bin");

        foreach (var pt in _pivotTables)
        {
            names.Add($"xl/pivotTables/pivotTable{pt.PivotTableIndex}.xml");
            names.Add($"xl/pivotCache/pivotCacheDefinition{pt.CacheId + 1}.xml");
            names.Add($"xl/pivotCache/pivotCacheRecords{pt.CacheId + 1}.xml");
        }

        return names;
    }

    private void ReadPictures(ZipArchive archive)
    {
        _pictures.Clear();
        foreach (var entry in archive.Entries
            .Where(entry => entry.FullName.StartsWith("xl/media/image", StringComparison.Ordinal))
            .OrderBy(entry => entry.FullName, StringComparer.Ordinal))
        {
            using var entryStream = entry.Open();
            using var memory = new MemoryStream();
            entryStream.CopyTo(memory);
            _pictures.Add(new XSSFPictureData(memory.ToArray(), GetPictureTypeFromExtension(Path.GetExtension(entry.FullName)), _pictures.Count + 1));
        }
    }

    private static List<XSSFRichTextString> ReadSharedStrings(ZipArchive archive)
    {
        var entry = archive.GetEntry("xl/sharedStrings.xml");
        if (entry is null)
        {
            return new List<XSSFRichTextString>();
        }

        using var stream = entry.Open();
        using var reader = XmlReader.Create(stream, new XmlReaderSettings { IgnoreWhitespace = false });
        var sharedStrings = new List<XSSFRichTextString>();

        while (reader.Read())
        {
            if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "si")
            {
                sharedStrings.Add(ReadRichSi(reader));
            }
        }

        return sharedStrings;
    }

    /// <summary>
    /// Reads a single &lt;si&gt; element, returning an XSSFRichTextString.
    /// If the si contains &lt;r&gt; elements, constructs a rich text with per-run formatting.
    /// Otherwise falls back to concatenating all &lt;t&gt; descendants (plain text).
    /// Ported from POI's SharedStringsTable parsing.
    /// </summary>
    private static XSSFRichTextString ReadRichSi(XmlReader siReader)
    {
        if (siReader.IsEmptyElement)
            return new XSSFRichTextString(string.Empty);

        using var subtree = siReader.ReadSubtree();
        var runs = new List<XSSFRichTextString.TextRun>();
        var plainText = new StringBuilder();
        bool hasRuns = false;

        while (subtree.Read())
        {
            if (subtree.NodeType == XmlNodeType.Element && subtree.LocalName == "r")
            {
                hasRuns = true;
                var run = ReadTextRun(subtree);
                runs.Add(run);
                continue;
            }

            if (subtree.NodeType == XmlNodeType.Element && subtree.LocalName == "t" && !hasRuns)
            {
                plainText.Append(subtree.ReadElementContentAsString());
                continue;
            }
        }

        if (hasRuns)
            return new XSSFRichTextString(runs);

        return new XSSFRichTextString(plainText.ToString());
    }

    /// <summary>Reads a single &lt;r&gt; element: optional &lt;rPr&gt; + mandatory &lt;t&gt;.</summary>
    private static XSSFRichTextString.TextRun ReadTextRun(XmlReader runReader)
    {
        var run = new XSSFRichTextString.TextRun();

        if (runReader.IsEmptyElement)
            return run;

        using var subtree = runReader.ReadSubtree();
        bool inRPr = false;

        while (subtree.Read())
        {
            if (subtree.NodeType == XmlNodeType.Element && subtree.LocalName == "rPr")
            {
                inRPr = true;
                continue;
            }
            if (subtree.NodeType == XmlNodeType.EndElement && subtree.LocalName == "rPr")
            {
                inRPr = false;
                continue;
            }

            if (inRPr)
            {
                // Element presence => formatting flags
                if (subtree.NodeType == XmlNodeType.Element)
                {
                    switch (subtree.LocalName)
                    {
                        case "b":
                        case "bChar":
                            run.Bold = true;
                            break;
                        case "i":
                        case "iChar":
                            run.Italic = true;
                            break;
                        case "u":
                            run.Underline = true;
                            break;
                        case "strike":
                            run.Strikethrough = true;
                            break;
                        case "sz":
                            var szVal = subtree.GetAttribute("val");
                            if (szVal is not null && double.TryParse(szVal,
                                NumberStyles.Integer, CultureInfo.InvariantCulture, out var sz))
                                run.FontSize = sz / 100.0; // OOXML sz is in hundredths of a point
                            break;
                        case "rFont":
                            var tf = subtree.GetAttribute("val") ?? subtree.GetAttribute("ascii");
                            if (tf is not null) run.FontName = tf;
                            break;
                        case "color":
                            var cv = subtree.GetAttribute("rgb") ?? subtree.GetAttribute("val");
                            if (cv is not null) run.Color = cv;
                            break;
                    }
                }
                continue;
            }

            // Text content
            if (subtree.NodeType == XmlNodeType.Element && subtree.LocalName == "t")
            {
                run.Text = subtree.ReadElementContentAsString();
            }
        }

        return run;
    }

    private static string ReadTextDescendants(XmlReader reader)
    {
        var text = new StringBuilder();

        if (!reader.IsEmptyElement)
        {
            using var subtree = reader.ReadSubtree();
            while (subtree.Read())
            {
                if (subtree.NodeType == XmlNodeType.Element && subtree.LocalName == "t")
                {
                    text.Append(subtree.ReadElementContentAsString());
                }
            }
        }

        return text.ToString();
    }

    private static List<SheetInfo> ReadWorkbookSheets(ZipArchive archive)
    {
        var workbookEntry = archive.GetEntry("xl/workbook.xml")
            ?? throw new InvalidDataException("The xlsx package is missing xl/workbook.xml.");
        var relationshipsEntry = archive.GetEntry("xl/_rels/workbook.xml.rels")
            ?? throw new InvalidDataException("The xlsx package is missing xl/_rels/workbook.xml.rels.");

        using var workbookStream = workbookEntry.Open();
        using var relationshipsStream = relationshipsEntry.Open();
        var targetsById = new Dictionary<string, string>(StringComparer.Ordinal);
        using (var relationshipReader = XmlReader.Create(relationshipsStream, new XmlReaderSettings { IgnoreWhitespace = false }))
        {
            while (relationshipReader.Read())
            {
                if (relationshipReader.NodeType != XmlNodeType.Element || relationshipReader.LocalName != "Relationship")
                {
                    continue;
                }

                var id = relationshipReader.GetAttribute("Id");
                var target = relationshipReader.GetAttribute("Target");
                var type = relationshipReader.GetAttribute("Type");
                if (id is not null
                    && target is not null
                    && string.Equals(type, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet", StringComparison.Ordinal))
                {
                    targetsById[id] = NormalizeWorkbookRelationshipTarget(target);
                }
            }
        }

        var sheets = new List<SheetInfo>();
        using (var workbookReader = XmlReader.Create(workbookStream, new XmlReaderSettings { IgnoreWhitespace = false }))
        {
            while (workbookReader.Read())
            {
                if (workbookReader.NodeType != XmlNodeType.Element || workbookReader.LocalName != "sheet")
                {
                    continue;
                }

                var name = workbookReader.GetAttribute("name")
                    ?? throw new InvalidDataException("A workbook sheet is missing its name attribute.");
                var relationshipId = GetAttributeByLocalName(workbookReader, "id")
                    ?? throw new InvalidDataException($"Sheet '{name}' is missing its relationship id.");
                if (!targetsById.TryGetValue(relationshipId, out var target))
                {
                    throw new InvalidDataException($"Sheet '{name}' references missing relationship '{relationshipId}'.");
                }

                sheets.Add(new SheetInfo(name, target));
            }
        }

        return sheets;
    }

    private static void ReadWorksheet(ZipArchive archive, string partName, XSSFSheet sheet, IReadOnlyList<XSSFRichTextString> sharedStrings)
    {
        var entry = archive.GetEntry(partName)
            ?? throw new InvalidDataException($"The xlsx package is missing {partName}.");

        // Read sheet rels for hyperlinks
        var hyperlinkRels = ReadSheetHyperlinkRelationships(archive, partName);

        using var stream = entry.Open();
        using var reader = XmlReader.Create(stream, new XmlReaderSettings { IgnoreWhitespace = false });
        XSSFRow? currentRow = null;
        XSSFCell? currentCell = null;
        string? currentCellTypeAttr = null; // raw "t" attribute
        bool currentCellIsFormula = false;  // true when <f> element seen inside current <c>
        bool currentCellHasValue = false;
        string? currentFormulaText = null;
        bool inFormulaElement = false;
        var formulaText = new StringBuilder();
        var inlineText = new StringBuilder();
        var nextRowIndex = 0;
        var nextColumnIndex = 0;
        bool inCols = false;

        while (reader.Read())
        {
            if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "sheetViews")
            {
                using var subtree = reader.ReadSubtree();
                while (subtree.Read())
                {
                    if (subtree.NodeType == XmlNodeType.Element && subtree.LocalName == "pane")
                    {
                        var xSplitAttr = subtree.GetAttribute("xSplit");
                        var ySplitAttr = subtree.GetAttribute("ySplit");
                        var xS = xSplitAttr is not null && int.TryParse(xSplitAttr, out var xs) ? xs : 0;
                        var yS = ySplitAttr is not null && int.TryParse(ySplitAttr, out var ys) ? ys : 0;
                        sheet.createFreezePane(xS, yS);
                    }
                }
                continue;
            }
            // Parse sheet protection
            if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "sheetProtection")
            {
                sheet.protectSheet(true);
                continue;
            }
            // Parse auto filter
            if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "autoFilter")
            {
                var refAttr = reader.GetAttribute("ref");
                if (refAttr is not null)
                {
                    try
                    {
                        var range = CellRangeAddress.Parse(refAttr);
                        sheet.setAutoFilter(range);
                    }
                    catch (ArgumentException)
                    {
                        // Ignore malformed range references
                    }
                }
                continue;
            }
            if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "cols")
            {
                inCols = !reader.IsEmptyElement;
                continue;
            }
            if (reader.NodeType == XmlNodeType.EndElement && reader.LocalName == "cols")
            {
                inCols = false;
                continue;
            }

            if (inCols && reader.NodeType == XmlNodeType.Element && reader.LocalName == "col")
            {
                var minAttr = reader.GetAttribute("min");
                var widthAttr = reader.GetAttribute("width");
                var hiddenAttr = reader.GetAttribute("hidden");
                if (minAttr is not null
                    && int.TryParse(minAttr, NumberStyles.Integer, CultureInfo.InvariantCulture, out var colNum))
                {
                    var colIndex = colNum - 1;
                    if (widthAttr is not null
                        && double.TryParse(widthAttr, NumberStyles.Float, CultureInfo.InvariantCulture, out var width))
                    {
                        sheet.setColumnWidth(colIndex, (int)(width * 256));
                    }
                    if (hiddenAttr == "true" || hiddenAttr == "1")
                        sheet.setColumnHidden(colIndex, true);
                }
                continue;
            }

            if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "row")
            {
                var rowIndex = ParseOneBasedAttribute(reader.GetAttribute("r"), nextRowIndex + 1);
                currentRow = sheet.createRow(rowIndex);
                nextRowIndex = rowIndex + 1;
                nextColumnIndex = 0;
                // Row height
                var htAttr = reader.GetAttribute("ht");
                if (htAttr is not null && double.TryParse(htAttr, NumberStyles.Float, CultureInfo.InvariantCulture, out var ht))
                {
                    currentRow.setHeight((float)ht);
                }
                // Hidden row
                var hiddenAttr = reader.GetAttribute("hidden");
                if (hiddenAttr == "true" || hiddenAttr == "1")
                    currentRow.setHidden(true);
                continue;
            }

            if (reader.NodeType == XmlNodeType.EndElement && reader.LocalName == "row")
            {
                currentRow = null;
                continue;
            }

            if (currentRow is not null && reader.NodeType == XmlNodeType.Element && reader.LocalName == "c")
            {
                var reference = reader.GetAttribute("r");
                var columnIndex = reference is null ? nextColumnIndex : ParseColumnIndex(reference);
                currentCell = currentRow.createCell(columnIndex);
                currentCellTypeAttr = reader.GetAttribute("t");
                currentCellIsFormula = false;
                currentCellHasValue = false;
                currentFormulaText = null;
                inFormulaElement = false;
                formulaText.Clear();
                var styleIndexText = reader.GetAttribute("s");
                if (styleIndexText is not null)
                {
                    currentCell.SetCellStyleIndex(int.Parse(styleIndexText, CultureInfo.InvariantCulture));
                }
                inlineText.Clear();
                nextColumnIndex = columnIndex + 1;

                if (reader.IsEmptyElement)
                {
                    // Empty element <c/> → blank cell, leave as Blank
                    currentCell = null;
                    currentCellTypeAttr = null;
                    currentCellIsFormula = false;
                    currentCellHasValue = false;
                    currentFormulaText = null;
                    inFormulaElement = false;
                    formulaText.Clear();
                }
                continue;
            }

            if (inFormulaElement)
            {
                if (reader.NodeType == XmlNodeType.Text || reader.NodeType == XmlNodeType.CDATA)
                    formulaText.Append(reader.Value);
                else if (reader.NodeType == XmlNodeType.EndElement && reader.LocalName == "f")
                {
                    currentFormulaText = formulaText.ToString();
                    inFormulaElement = false;
                }
                continue;
            }

            if (currentCell is not null && reader.NodeType == XmlNodeType.Element && reader.LocalName == "f")
            {
                currentCellIsFormula = true;
                formulaText.Clear();
                if (reader.IsEmptyElement)
                    currentFormulaText = string.Empty;
                else
                    inFormulaElement = true;
                continue;
            }

            if (currentCell is not null && reader.NodeType == XmlNodeType.Element && reader.LocalName == "v")
            {
                var rawValue = reader.ReadElementContentAsString();
                currentCellHasValue = true;
                ApplyCellValue(currentCell, currentCellTypeAttr, currentCellIsFormula, currentFormulaText, rawValue, sharedStrings);
                continue;
            }

            // Inline string segment
            if (currentCell is not null && currentCellTypeAttr == "inlineStr"
                && reader.NodeType == XmlNodeType.Element && reader.LocalName == "t")
            {
                inlineText.Append(reader.ReadElementContentAsString());
                continue;
            }

            if (currentCell is not null && reader.NodeType == XmlNodeType.EndElement && reader.LocalName == "c")
            {
                if (currentCellTypeAttr == "inlineStr")
                {
                    if (currentCellIsFormula)
                        currentCell.SetFormulaFromXml(currentFormulaText, CellType.String, inlineText.ToString(), hasCachedValue: true);
                    else
                        currentCell.SetValueFromXml(CellType.String, inlineText.ToString());
                }
                else if (currentCellIsFormula && !currentCellHasValue)
                {
                    var cachedType = currentCellTypeAttr switch
                    {
                        "b" => CellType.Boolean,
                        "e" => CellType.Error,
                        "str" => CellType.String,
                        "s" => CellType.String,
                        _ => CellType.Numeric
                    };
                    currentCell.SetFormulaFromXml(currentFormulaText, cachedType, null, hasCachedValue: false);
                }

                currentCell = null;
                currentCellTypeAttr = null;
                currentCellIsFormula = false;
                currentCellHasValue = false;
                currentFormulaText = null;
                inFormulaElement = false;
                formulaText.Clear();
            }

            // Parse merge cells
            if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "mergeCell")
            {
                var refAttr = reader.GetAttribute("ref");
                if (refAttr is not null)
                {
                    var region = CellRangeAddress.Parse(refAttr);
                    sheet.addMergedRegion(region);
                }
                continue;
            }

            // Parse page margins
            if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "pageMargins")
            {
                sheet.PageMarginBottom = ParseDoubleAttr(reader, "bottom", 0.75);
                sheet.PageMarginFooter = ParseDoubleAttr(reader, "footer", 0.3);
                sheet.PageMarginHeader = ParseDoubleAttr(reader, "header", 0.3);
                sheet.PageMarginLeft = ParseDoubleAttr(reader, "left", 0.7);
                sheet.PageMarginRight = ParseDoubleAttr(reader, "right", 0.7);
                sheet.PageMarginTop = ParseDoubleAttr(reader, "top", 0.75);
                continue;
            }

            // Parse page setup
            if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "pageSetup")
            {
                sheet.PageOrientation = reader.GetAttribute("orientation");
                var ftw = reader.GetAttribute("fitToWidth");
                if (ftw is not null && int.TryParse(ftw, NumberStyles.Integer, CultureInfo.InvariantCulture, out var ftwVal))
                    sheet.FitToWidth = ftwVal;
                var fth = reader.GetAttribute("fitToHeight");
                if (fth is not null && int.TryParse(fth, NumberStyles.Integer, CultureInfo.InvariantCulture, out var fthVal))
                    sheet.FitToHeight = fthVal;
                var ps = reader.GetAttribute("paperSize");
                if (ps is not null && int.TryParse(ps, NumberStyles.Integer, CultureInfo.InvariantCulture, out var psVal))
                    sheet.PaperSize = psVal;
                continue;
            }

            // Parse header/footer
            if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "oddHeader")
            {
                sheet.HeaderCenter = reader.ReadString();
                continue;
            }
            if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "oddFooter")
            {
                sheet.FooterCenter = reader.ReadString();
                continue;
            }

            // Parse hyperlinks
            if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "hyperlink")
            {
                var refAttr = reader.GetAttribute("ref");
                var relId = GetAttributeByLocalName(reader, "id");
                if (refAttr is not null && relId is not null && hyperlinkRels.TryGetValue(relId, out var hyperlinkInfo))
                {
                    var (rowNum, colNum) = ParseCellRef(refAttr);
                    var hyperlink = new XSSFHyperlink(hyperlinkInfo.IsExternal ? HyperlinkType.Url : HyperlinkType.Document)
                    {
                        Address = hyperlinkInfo.Target,
                        CellRef = refAttr,
                        IsExternal = hyperlinkInfo.IsExternal,
                        RelationshipId = relId
                    };
                    sheet.AddHyperlink(hyperlink);
                    var existingRow = sheet.getRow(rowNum);
                    if (existingRow is XSSFRow existingXSSFRow)
                    {
                        var existingCell = existingXSSFRow.getCell(colNum);
                        if (existingCell is XSSFCell targetCell)
                        {
                            targetCell.setHyperlink(hyperlink);
                        }
                    }
                }
                continue;
            }

            // Parse data validation
            if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "dataValidations")
            {
                continue;
            }
            if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "dataValidation")
            {
                var dv = new XSSFDataValidation();
                var typeAttr = reader.GetAttribute("type");
                if (typeAttr is not null)
                    dv.Type = DataValidationTypeFromName(typeAttr);
                var opAttr = reader.GetAttribute("operator");
                if (opAttr is not null)
                    dv.Operator = DataValidationOperatorFromName(opAttr);
                dv.Sqref = reader.GetAttribute("sqref") ?? string.Empty;
                dv.AllowBlank = ParseBooleanAttribute(reader.GetAttribute("allowBlank"), defaultValue: true);
                dv.ShowInputMessage = ParseBooleanAttribute(reader.GetAttribute("showInputMessage"), defaultValue: true);
                dv.ShowErrorMessage = ParseBooleanAttribute(reader.GetAttribute("showErrorMessage"), defaultValue: true);
                dv.ShowDropDown = ParseBooleanAttribute(reader.GetAttribute("showDropDown"), defaultValue: true);
                dv.ErrorStyle = reader.GetAttribute("errorStyle");
                dv.ErrorTitle = reader.GetAttribute("errorTitle");
                dv.ErrorMessage = reader.GetAttribute("error");
                dv.PromptTitle = reader.GetAttribute("promptTitle");
                dv.PromptMessage = reader.GetAttribute("prompt");
                if (!reader.IsEmptyElement)
                {
                    var dvDepth = reader.Depth;
                    while (reader.Read() && reader.Depth > dvDepth)
                    {
                        if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "formula1" && !reader.IsEmptyElement)
                            dv.Formula1 = reader.ReadString();
                        if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "formula2" && !reader.IsEmptyElement)
                            dv.Formula2 = reader.ReadString();
                    }
                }
                sheet.AddDataValidation(dv);
                continue;
            }

            // Parse conditional formatting
            if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "conditionalFormatting")
            {
                var cf = new XSSFConditionalFormatting();
                cf.Sqref = reader.GetAttribute("sqref") ?? string.Empty;
                if (!reader.IsEmptyElement)
                {
                    var cfDepth = reader.Depth;
                    while (reader.Read() && reader.Depth > cfDepth)
                    {
                        if (reader.NodeType != XmlNodeType.Element) continue;
                        if (reader.LocalName == "cfRule")
                        {
                            var rule = new XSSFCFRule();
                            var rTypeAttr = reader.GetAttribute("type");
                            if (rTypeAttr is not null)
                                rule.Type = CfTypeFromName(rTypeAttr);
                            var priorityAttr = reader.GetAttribute("priority");
                            if (priorityAttr is not null && int.TryParse(priorityAttr, NumberStyles.Integer, CultureInfo.InvariantCulture, out var pri))
                                rule.Priority = pri;
                            rule.Operator = reader.GetAttribute("operator");
                            rule.Text = reader.GetAttribute("text");
                            var dxfAttr = reader.GetAttribute("dxfId");
                            if (dxfAttr is not null && int.TryParse(dxfAttr, NumberStyles.Integer, CultureInfo.InvariantCulture, out var dxfId))
                                rule.DxfId = dxfId;
                            if (!reader.IsEmptyElement)
                            {
                                var ruleDepth = reader.Depth;
                                while (reader.Read() && reader.Depth > ruleDepth)
                                {
                                    if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "formula" && !reader.IsEmptyElement)
                                    {
                                        rule.Formulas.Add(reader.ReadString());
                                    }
                                }
                            }
                            cf.Rules.Add(rule);
                        }
                    }
                }
                sheet.AddConditionalFormatting(cf);
                continue;
            }
        }
    }

    /// <summary>
    /// Maps the raw OOXML t attribute to a CellType.
    /// Ported from XSSFCell.getBaseCellType() / STCellType constants.
    /// </summary>
    private static CellType ParseCellTypeAttr(string? t, string? rawValue)
    {
        return t switch
        {
            "b"         => CellType.Boolean,
            "e"         => CellType.Error,
            "s"         => CellType.String,   // shared string
            "str"       => CellType.String,   // formula cached string
            "inlineStr" => CellType.String,   // inline string (handled separately)
            null or "n" => rawValue is null ? CellType.Blank : CellType.Numeric,
            _ => CellType.Numeric             // unknown t → treat as numeric
        };
    }

    private static void ApplyCellValue(
        XSSFCell cell, string? cellTypeAttr, bool isFormula, string? formulaText,
        string rawValue, IReadOnlyList<XSSFRichTextString> sharedStrings)
    {
        // Resolve the base cell type from the t attribute
        // For numeric types: if rawValue is empty/null → Blank (no <v> set)
        var baseType = ParseCellTypeAttr(cellTypeAttr, rawValue.Length > 0 ? rawValue : null);

        if (baseType == CellType.String && cellTypeAttr == "s")
        {
            // Shared string: rawValue is the SST index
            var idx = int.Parse(rawValue, CultureInfo.InvariantCulture);
            if ((uint)idx >= (uint)sharedStrings.Count)
                throw new InvalidDataException($"Shared string index {idx} is outside the table.");
            var rts = sharedStrings[idx];
            rawValue = rts.getString();
            cell.SetRichTextStringFromSst(rts);
        }

        if (isFormula)
            cell.SetFormulaFromXml(formulaText, baseType, rawValue, hasCachedValue: true);
        else
            cell.SetValueFromXml(baseType, rawValue);
    }

    private void ReadSheetDrawing(ZipArchive archive, string sheetPartName, XSSFSheet sheet)
    {
        // e.g. "xl/worksheets/sheet1.xml" → "xl/worksheets/_rels/sheet1.xml.rels"
        var dir = Path.GetDirectoryName(sheetPartName)?.Replace('\\', '/') ?? string.Empty;
        var file = Path.GetFileName(sheetPartName);
        var relsPath = $"{dir}/_rels/{file}.rels";

        var relsEntry = archive.GetEntry(relsPath);
        if (relsEntry is null) return;

        string? drawingTarget = null;
        using (var relsStream = relsEntry.Open())
        using (var relsReader = XmlReader.Create(relsStream, new XmlReaderSettings { IgnoreWhitespace = false }))
        {
            while (relsReader.Read())
            {
                if (relsReader.NodeType != XmlNodeType.Element || relsReader.LocalName != "Relationship") continue;
                if (!string.Equals(relsReader.GetAttribute("Type"),
                        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing",
                        StringComparison.Ordinal))
                    continue;
                drawingTarget = relsReader.GetAttribute("Target");
                break;
            }
        }

        if (drawingTarget is null) return;

        // Target is relative to xl/worksheets/, so "../drawings/drawing1.xml" → "xl/drawings/drawing1.xml"
        var drawingPath = NormalizeRelativePath(dir + "/" + drawingTarget);
        var drawingDir = Path.GetDirectoryName(drawingPath)?.Replace('\\', '/') ?? string.Empty;
        var drawingFile = Path.GetFileName(drawingPath);
        var drawingRelsPath = $"{drawingDir}/_rels/{drawingFile}.rels";

        ReadDrawing(archive, drawingPath, drawingRelsPath, sheet);
    }

    private void ReadDrawing(ZipArchive archive, string drawingPath, string drawingRelsPath, XSSFSheet sheet)
    {
        var drawingEntry = archive.GetEntry(drawingPath);
        if (drawingEntry is null) return;

        // Map relationship id → 0-based picture index in _pictures
        var pictureIndexByRelId = new Dictionary<string, int>(StringComparer.Ordinal);
        var relsEntry = archive.GetEntry(drawingRelsPath);
        if (relsEntry is not null)
        {
            using var relsStream = relsEntry.Open();
            using var relsReader = XmlReader.Create(relsStream, new XmlReaderSettings { IgnoreWhitespace = false });
            while (relsReader.Read())
            {
                if (relsReader.NodeType != XmlNodeType.Element || relsReader.LocalName != "Relationship") continue;
                if (!string.Equals(relsReader.GetAttribute("Type"),
                        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
                        StringComparison.Ordinal))
                    continue;
                var relId = relsReader.GetAttribute("Id");
                var target = relsReader.GetAttribute("Target"); // e.g. "../media/image1.jpeg"
                if (relId is null || target is null) continue;
                var mediaFile = Path.GetFileNameWithoutExtension(target); // "image1"
                if (mediaFile.StartsWith("image", StringComparison.OrdinalIgnoreCase)
                    && int.TryParse(mediaFile["image".Length..], out var oneBasedIndex))
                {
                    pictureIndexByRelId[relId] = oneBasedIndex - 1;
                }
            }
        }

        var drawing = sheet.createDrawingPatriarch();

        // Parse twoCellAnchor elements to reconstruct XSSFPicture shapes
        using var drawingStream = drawingEntry.Open();
        using var reader = XmlReader.Create(drawingStream, new XmlReaderSettings { IgnoreWhitespace = false });

        XSSFClientAnchor? currentAnchor = null;
        string? currentRelId = null;
        int currentShapeId = 0;
        bool inFrom = false, inTo = false;
        int rotAttribute = 0;
        // from/to marker accumulators
        int fCol = 0, fDx = 0, fRow = 0, fDy = 0;
        int tCol = 0, tDx = 0, tRow = 0, tDy = 0;
        string? markerField = null;
        int markerValue = 0;
        bool inMarkerElement = false;

        while (reader.Read())
        {
            if (reader.NodeType == XmlNodeType.Element)
            {
                switch (reader.LocalName)
                {
                    case "twoCellAnchor":
                        currentAnchor = null;
                        currentRelId = null;
                        currentShapeId = 0;
                        rotAttribute = 0;
                        fCol = fDx = fRow = fDy = 0;
                        tCol = tDx = tRow = tDy = 0;
                        break;
                    case "from":
                        inFrom = true; inTo = false;
                        break;
                    case "to":
                        inTo = true; inFrom = false;
                        break;
                    case "col":
                    case "colOff":
                    case "row":
                    case "rowOff":
                        markerField = reader.LocalName;
                        inMarkerElement = !reader.IsEmptyElement;
                        break;
                    case "cNvPr" when currentAnchor is null:
                        var idAttr = reader.GetAttribute("id");
                        if (idAttr is not null) int.TryParse(idAttr, out currentShapeId);
                        break;
                    case "blip":
                        var embedAttr = reader.GetAttribute("embed");
                        if (embedAttr is null)
                        {
                            // try namespace-qualified attribute
                            embedAttr = GetAttributeByLocalName(reader, "embed");
                        }
                        currentRelId = embedAttr;
                        break;
                    case "xfrm":
                        var rotAttr = reader.GetAttribute("rot");
                        rotAttribute = rotAttr is not null ? int.Parse(rotAttr, CultureInfo.InvariantCulture) : 0;
                        break;
                }
            }
            else if (reader.NodeType == XmlNodeType.Text && inMarkerElement && markerField is not null)
            {
                if (int.TryParse(reader.Value, out markerValue))
                {
                    if (inFrom)
                    {
                        switch (markerField)
                        {
                            case "col": fCol = markerValue; break;
                            case "colOff": fDx = markerValue; break;
                            case "row": fRow = markerValue; break;
                            case "rowOff": fDy = markerValue; break;
                        }
                    }
                    else if (inTo)
                    {
                        switch (markerField)
                        {
                            case "col": tCol = markerValue; break;
                            case "colOff": tDx = markerValue; break;
                            case "row": tRow = markerValue; break;
                            case "rowOff": tDy = markerValue; break;
                        }
                    }
                }
            }
            else if (reader.NodeType == XmlNodeType.EndElement)
            {
                switch (reader.LocalName)
                {
                    case "col":
                    case "colOff":
                    case "row":
                    case "rowOff":
                        inMarkerElement = false;
                        markerField = null;
                        break;
                    case "from":
                        inFrom = false;
                        currentAnchor = new XSSFClientAnchor(fDx, fDy, tDx, tDy, fCol, fRow, tCol, tRow);
                        break;
                    case "to":
                        inTo = false;
                        break;
                    case "twoCellAnchor":
                        if (currentAnchor is not null && currentRelId is not null
                            && pictureIndexByRelId.TryGetValue(currentRelId, out var picIdx))
                        {
                            var picture = drawing.createPicture(currentAnchor, picIdx);
                            picture.SetRotationAttribute(rotAttribute);
                        }
                        break;
                }
            }
        }
    }

    private static string NormalizeRelativePath(string path)
    {
        var parts = path.Split('/');
        var stack = new System.Collections.Generic.Stack<string>();
        foreach (var part in parts)
        {
            if (part == "..")
            {
                if (stack.Count > 0) stack.Pop();
            }
            else if (part != "." && part != string.Empty)
            {
                stack.Push(part);
            }
        }
        return string.Join("/", stack.Reverse());
    }

    private static Dictionary<string, (string Target, bool IsExternal)> ReadSheetHyperlinkRelationships(ZipArchive archive, string sheetPartName)
    {
        var result = new Dictionary<string, (string, bool)>(StringComparer.Ordinal);
        var dir = Path.GetDirectoryName(sheetPartName)?.Replace('\\', '/') ?? string.Empty;
        var file = Path.GetFileName(sheetPartName);
        var relsPath = $"{dir}/_rels/{file}.rels";
        var relsEntry = archive.GetEntry(relsPath);
        if (relsEntry is null) return result;
        using var relsStream = relsEntry.Open();
        using var relsReader = XmlReader.Create(relsStream, new XmlReaderSettings { IgnoreWhitespace = false });
        while (relsReader.Read())
        {
            if (relsReader.NodeType != XmlNodeType.Element || relsReader.LocalName != "Relationship") continue;
            var relType = relsReader.GetAttribute("Type");
            if (relType is null || !relType.EndsWith("/relationships/hyperlink", StringComparison.Ordinal)) continue;
            var relId = relsReader.GetAttribute("Id");
            var target = relsReader.GetAttribute("Target");
            var targetMode = relsReader.GetAttribute("TargetMode");
            if (relId is not null && target is not null)
            {
                result[relId] = (target, string.Equals(targetMode, "External", StringComparison.Ordinal));
            }
        }
        return result;
    }

    private static (int row, int col) ParseCellRef(string reference)
    {
        var clean = reference.TrimStart('$');
        int i = 0;
        while (i < clean.Length && char.IsLetter(clean[i])) i++;
        var colPart = clean.Substring(0, i);
        var rowPart = clean.Substring(i);
        int col = 0;
        foreach (var ch in colPart)
        {
            col = col * 26 + (char.ToUpperInvariant(ch) - 'A' + 1);
        }
        return (int.Parse(rowPart, CultureInfo.InvariantCulture) - 1, col - 1);
    }

    private void ReadStyles(ZipArchive archive)
    {
        var entry = archive.GetEntry("xl/styles.xml");
        if (entry is null)
        {
            return;
        }

        InitializeDefaultStyles();
        using var stream = entry.Open();
        using var reader = XmlReader.Create(stream, new XmlReaderSettings { IgnoreWhitespace = false });

        while (reader.Read())
        {
            if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "numFmt")
            {
                var id = int.Parse(reader.GetAttribute("numFmtId") ?? "0", CultureInfo.InvariantCulture);
                var code = reader.GetAttribute("formatCode") ?? string.Empty;
                _customNumberFormats[id] = code;
                _customNumberFormatIndexes[code] = id;
                continue;
            }

            if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "fonts")
            {
                ReadFonts(reader);
                continue;
            }

            if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "fills")
            {
                ReadFills(reader);
                continue;
            }

            if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "borders")
            {
                ReadBorders(reader);
                continue;
            }

            if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "cellXfs")
            {
                ReadCellXfs(reader);
            }
        }
    }

    private void ReadFonts(XmlReader reader)
    {
        _fonts.Clear();
        if (reader.IsEmptyElement)
        {
            _fonts.Add(new XSSFFont(0));
            return;
        }

        using var subtree = reader.ReadSubtree();
        XSSFFont? font = null;
        while (subtree.Read())
        {
            if (subtree.NodeType == XmlNodeType.Element && subtree.LocalName == "font")
            {
                font = new XSSFFont(_fonts.Count);
                _fonts.Add(font);
                continue;
            }

            if (font is null || subtree.NodeType != XmlNodeType.Element)
            {
                continue;
            }

            switch (subtree.LocalName)
            {
                case "b":
                    font.setBold(ParseBooleanAttribute(subtree.GetAttribute("val"), defaultValue: true));
                    break;
                case "i":
                    font.setItalic(ParseBooleanAttribute(subtree.GetAttribute("val"), defaultValue: true));
                    break;
                case "strike":
                    font.setStrikeout(ParseBooleanAttribute(subtree.GetAttribute("val"), defaultValue: true));
                    break;
                case "u":
                    font.setUnderline(1);
                    break;
                case "color":
                    if (short.TryParse(subtree.GetAttribute("indexed"), NumberStyles.Integer, CultureInfo.InvariantCulture, out var color))
                    {
                        font.setColor(color);
                    }
                    break;
                case "sz":
                    if (double.TryParse(subtree.GetAttribute("val"), NumberStyles.Float, CultureInfo.InvariantCulture, out var size))
                    {
                        font.setFontHeightInPoints((short)size);
                    }
                    break;
                case "name":
                    font.setFontName(subtree.GetAttribute("val"));
                    break;
            }
        }

        if (_fonts.Count == 0)
        {
            _fonts.Add(new XSSFFont(0));
        }
    }

    private void ReadFills(XmlReader reader)
    {
        _fills.Clear();
        // Always add default fills at index 0 and 1 (none and darkGray)
        _fills.Add(null);
        _fills.Add(null);

        if (reader.IsEmptyElement)
        {
            return;
        }

        using var subtree = reader.ReadSubtree();
        var fillIndex = 0;
        while (subtree.Read())
        {
            if (subtree.NodeType != XmlNodeType.Element || subtree.LocalName != "fill")
            {
                continue;
            }

            if (fillIndex < 2)
            {
                // Skip built-in default fills
                fillIndex++;
                continue;
            }

            // Read patternFill child
            FillPatternType pattern = FillPatternType.NoFill;
            short? fgColor = null;

            if (!subtree.IsEmptyElement)
            {
                using var fillSubtree = subtree.ReadSubtree();
                while (fillSubtree.Read())
                {
                    if (fillSubtree.NodeType != XmlNodeType.Element)
                    {
                        continue;
                    }

                    if (fillSubtree.LocalName == "patternFill")
                    {
                        var patternAttr = fillSubtree.GetAttribute("patternType");
                        if (patternAttr is not null)
                        {
                            int patternIdx = patternAttr switch
                            {
                                "none" => 1,
                                "solid" => 2,
                                "mediumGray" => 3,
                                "darkGray" => 4,
                                "lightGray" => 5,
                                "darkHorizontal" => 6,
                                "darkVertical" => 7,
                                "darkDown" => 8,
                                "darkUp" => 9,
                                "darkGrid" => 10,
                                "darkTrellis" => 11,
                                "lightHorizontal" => 12,
                                "lightVertical" => 13,
                                "lightDown" => 14,
                                "lightUp" => 15,
                                "lightGrid" => 16,
                                "lightTrellis" => 17,
                                "gray125" => 18,
                                "gray0625" => 19,
                                _ => 1
                            };
                            pattern = (FillPatternType)(patternIdx - 1);
                        }

                        // Read fgColor child
                        using var patternSubtree = fillSubtree.ReadSubtree();
                        while (patternSubtree.Read())
                        {
                            if (patternSubtree.NodeType == XmlNodeType.Element
                                && patternSubtree.LocalName == "fgColor")
                            {
                                var indexed = patternSubtree.GetAttribute("indexed");
                                if (indexed is not null
                                    && short.TryParse(indexed, NumberStyles.Integer, CultureInfo.InvariantCulture, out var c))
                                {
                                    fgColor = c;
                                }
                            }
                        }
                    }
                }
            }

            // Store this fill as a reusable style object
            var fillStyle = new XSSFCellStyle(this, fillIndex);
            fillStyle.FillPattern = pattern;
            fillStyle.FillForegroundColor = fgColor;
            _fills.Add(fillStyle);

            fillIndex++;
        }
    }

    private void ReadBorders(XmlReader reader)
    {
        _borders.Clear();
        _borders.Add(null); // index 0: default border (none)

        if (reader.IsEmptyElement)
        {
            return;
        }

        using var subtree = reader.ReadSubtree();
        var borderIndex = 0;
        while (subtree.Read())
        {
            if (subtree.NodeType != XmlNodeType.Element || subtree.LocalName != "border")
            {
                continue;
            }

            if (borderIndex == 0)
            {
                // Skip the first default border (index 0 in XML)
                borderIndex++;
                continue;
            }

            BorderStyle left = BorderStyle.None;
            BorderStyle right = BorderStyle.None;
            BorderStyle top = BorderStyle.None;
            BorderStyle bottom = BorderStyle.None;

            if (!subtree.IsEmptyElement)
            {
                using var borderSubtree = subtree.ReadSubtree();
                while (borderSubtree.Read())
                {
                    if (borderSubtree.NodeType != XmlNodeType.Element)
                    {
                        continue;
                    }

                    var styleName = borderSubtree.GetAttribute("style");
                    var borderStyle = styleName is null
                        ? BorderStyle.None
                        : ParseBorderStyleName(styleName);

                    switch (borderSubtree.LocalName)
                    {
                        case "left":   left   = borderStyle; break;
                        case "right":  right  = borderStyle; break;
                        case "top":    top    = borderStyle; break;
                        case "bottom": bottom = borderStyle; break;
                    }
                }
            }

            // Store this border as a reusable style object
            var borderObj = new XSSFCellStyle(this, borderIndex);
            borderObj.BorderLeft = left;
            borderObj.BorderRight = right;
            borderObj.BorderTop = top;
            borderObj.BorderBottom = bottom;
            _borders.Add(borderObj);

            borderIndex++;
        }
    }

    private static BorderStyle ParseBorderStyleName(string name)
    {
        return name switch
        {
            "thin"              => BorderStyle.Thin,
            "medium"            => BorderStyle.Medium,
            "dashed"            => BorderStyle.Dashed,
            "dotted"            => BorderStyle.Dotted,
            "thick"             => BorderStyle.Thick,
            "double"            => BorderStyle.Double,
            "hair"              => BorderStyle.Hair,
            "mediumDashed"      => BorderStyle.MediumDashed,
            "dashDot"           => BorderStyle.DashDot,
            "mediumDashDot"     => BorderStyle.MediumDashDot,
            "dashDotDot"        => BorderStyle.DashDotDot,
            "mediumDashDotDot"  => BorderStyle.MediumDashDotDot,
            "slantDashDot"      => BorderStyle.SlantedDashDot,
            _                   => BorderStyle.None
        };
    }

    private static HorizontalAlignment ParseHorizontalAlignment(string name)
    {
        return name switch
        {
            "general"            => HorizontalAlignment.General,
            "left"               => HorizontalAlignment.Left,
            "center"             => HorizontalAlignment.Center,
            "right"              => HorizontalAlignment.Right,
            "fill"               => HorizontalAlignment.Fill,
            "justify"            => HorizontalAlignment.Justify,
            "centerContinuous"   => HorizontalAlignment.CenterSelection,
            "distributed"        => HorizontalAlignment.Distributed,
            _                    => HorizontalAlignment.General
        };
    }

    private static VerticalAlignment ParseVerticalAlignment(string name)
    {
        return name switch
        {
            "bottom"       => VerticalAlignment.Bottom,
            "center"       => VerticalAlignment.Center,
            "top"          => VerticalAlignment.Top,
            "justify"      => VerticalAlignment.Justify,
            "distributed"  => VerticalAlignment.Distributed,
            _              => VerticalAlignment.Bottom
        };
    }

    private void ReadCellXfs(XmlReader reader)
    {
        _cellStyles.Clear();
        if (reader.IsEmptyElement)
        {
            _cellStyles.Add(new XSSFCellStyle(this, 0));
            return;
        }

        using var subtree = reader.ReadSubtree();
        while (subtree.Read())
        {
            if (subtree.NodeType != XmlNodeType.Element || subtree.LocalName != "xf")
            {
                continue;
            }

            var style = new XSSFCellStyle(this, _cellStyles.Count);

            // numFmtId
            if (int.TryParse(subtree.GetAttribute("numFmtId"), NumberStyles.Integer, CultureInfo.InvariantCulture, out var numFmtId))
            {
                style.setDataFormat(numFmtId);
            }

            // fontId
            if (int.TryParse(subtree.GetAttribute("fontId"), NumberStyles.Integer, CultureInfo.InvariantCulture, out var fontId) && fontId != 0)
            {
                style.setFont(getFontAt(fontId));
            }

            // fillId — copy fill properties from the fills table
            if (int.TryParse(subtree.GetAttribute("fillId"), NumberStyles.Integer, CultureInfo.InvariantCulture, out var fillId))
            {
                style.FillId = fillId;
                if (fillId >= 0 && fillId < _fills.Count && _fills[fillId] is XSSFCellStyle srcFill)
                {
                    style.FillPattern = srcFill.FillPattern;
                    style.FillForegroundColor = srcFill.FillForegroundColor;
                }
            }
            style.ApplyFill = ParseBooleanAttribute(subtree.GetAttribute("applyFill"), defaultValue: true);

            // borderId — copy border properties from the borders table
            if (int.TryParse(subtree.GetAttribute("borderId"), NumberStyles.Integer, CultureInfo.InvariantCulture, out var borderId))
            {
                style.BorderId = borderId;
                if (borderId >= 0 && borderId < _borders.Count && _borders[borderId] is XSSFCellStyle srcBorder)
                {
                    style.BorderLeft = srcBorder.BorderLeft;
                    style.BorderRight = srcBorder.BorderRight;
                    style.BorderTop = srcBorder.BorderTop;
                    style.BorderBottom = srcBorder.BorderBottom;
                    style.ApplyBorder = true;
                }
            }
            if (!style.ApplyBorder)
            {
                // Also check the explicit attribute
                style.ApplyBorder = ParseBooleanAttribute(subtree.GetAttribute("applyBorder"), defaultValue: true);
            }

            // applyAlignment flag
            style.ApplyAlignment = ParseBooleanAttribute(subtree.GetAttribute("applyAlignment"), defaultValue: false);

            // Read alignment element from inside the xf
            if (!subtree.IsEmptyElement)
            {
                using var xfSubtree = subtree.ReadSubtree();
                while (xfSubtree.Read())
                {
                    if (xfSubtree.NodeType != XmlNodeType.Element || xfSubtree.LocalName != "alignment")
                    {
                        continue;
                    }

                    var hAttr = xfSubtree.GetAttribute("horizontal");
                    if (hAttr is not null)
                    {
                        style.AlignmentValue = ParseHorizontalAlignment(hAttr);
                        style.ApplyAlignment = true;
                    }

                    var vAttr = xfSubtree.GetAttribute("vertical");
                    if (vAttr is not null)
                    {
                        style.VerticalAlignmentValue = ParseVerticalAlignment(vAttr);
                        style.ApplyAlignment = true;
                    }

                    style.WrapTextEnabled = ParseBooleanAttribute(xfSubtree.GetAttribute("wrapText"), defaultValue: false);
                    if (style.WrapTextEnabled)
                    {
                        style.ApplyAlignment = true;
                    }

                    var indentAttr = xfSubtree.GetAttribute("indent");
                    if (indentAttr is not null
                        && short.TryParse(indentAttr, NumberStyles.Integer, CultureInfo.InvariantCulture, out var indent))
                    {
                        style.IndentLevel = indent;
                        style.ApplyAlignment = true;
                    }

                    var rotationAttr = xfSubtree.GetAttribute("textRotation");
                    if (rotationAttr is not null
                        && short.TryParse(rotationAttr, NumberStyles.Integer, CultureInfo.InvariantCulture, out var rotation))
                    {
                        style.TextRotation = rotation;
                        style.ApplyAlignment = true;
                    }
                }
            }

            _cellStyles.Add(style);
        }

        if (_cellStyles.Count == 0)
        {
            _cellStyles.Add(new XSSFCellStyle(this, 0));
        }
    }

    private static int ParseOneBasedAttribute(string? value, int defaultOneBasedValue)
    {
        var oneBasedValue = value is null
            ? defaultOneBasedValue
            : int.Parse(value, CultureInfo.InvariantCulture);
        return oneBasedValue - 1;
    }

    private static string NormalizeWorkbookRelationshipTarget(string target)
    {
        var partName = target.StartsWith("/", StringComparison.Ordinal)
            ? target.TrimStart('/')
            : "xl/" + target;
        var segments = new Stack<string>();

        foreach (var segment in partName.Split('/', StringSplitOptions.RemoveEmptyEntries))
        {
            if (segment == ".")
            {
                continue;
            }

            if (segment == "..")
            {
                if (segments.Count > 0)
                {
                    segments.Pop();
                }
                continue;
            }

            segments.Push(segment);
        }

        return string.Join("/", segments.Reverse());
    }

    private static string? GetAttributeByLocalName(XmlReader reader, string localName)
    {
        if (!reader.HasAttributes)
        {
            return null;
        }

        while (reader.MoveToNextAttribute())
        {
            if (reader.LocalName == localName)
            {
                var value = reader.Value;
                reader.MoveToElement();
                return value;
            }
        }

        reader.MoveToElement();
        return null;
    }

    private void WriteContentTypes(PoiXmlWriter writer)
    {
        writer.WriteStartElement("Types");
        writer.WriteAttributeString("xmlns", "http://schemas.openxmlformats.org/package/2006/content-types");
        WriteDefault(writer, "rels", "application/vnd.openxmlformats-package.relationships+xml");
        foreach (var pictureDefault in _pictures
            .GroupBy(picture => picture.Extension, StringComparer.Ordinal)
            .Select(group => group.First())
            .OrderBy(picture => picture.Extension, StringComparer.Ordinal))
        {
            WriteDefault(writer, pictureDefault.Extension, pictureDefault.ContentType);
        }
        WriteDefault(writer, "xml", "application/xml");
        WriteOverride(writer, "/docProps/app.xml", "application/vnd.openxmlformats-officedocument.extended-properties+xml");
        WriteOverride(writer, "/docProps/core.xml", "application/vnd.openxmlformats-package.core-properties+xml");
        // xlsm uses macroEnabled content type; xlsx uses the standard one.
        WriteOverride(writer, "/xl/workbook.xml", _vbaProjectBin != null ? ContentTypeXlsm : ContentTypeXlsx);
        if (_vbaProjectBin != null)
            WriteOverride(writer, "/xl/vbaProject.bin", ContentTypeVbaProject);
        WriteOverride(writer, "/xl/sharedStrings.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml");
        WriteOverride(writer, "/xl/styles.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml");
        foreach (var sheet in _sheets)
        {
            WriteOverride(writer, $"/xl/worksheets/sheet{sheet.SheetIndex}.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml");
            if (sheet.Drawing is not null)
            {
                WriteOverride(writer, $"/xl/drawings/drawing{sheet.Drawing.DrawingIndex}.xml", "application/vnd.openxmlformats-officedocument.drawing+xml");
            }
        }
        // Pivot table content types
        foreach (var pt in _pivotTables)
        {
            WriteOverride(writer, $"/xl/pivotTables/pivotTable{pt.PivotTableIndex}.xml",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml");
            WriteOverride(writer, $"/xl/pivotCache/pivotCacheDefinition{pt.CacheId + 1}.xml",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml");
            WriteOverride(writer, $"/xl/pivotCache/pivotCacheRecords{pt.CacheId + 1}.xml",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml");
        }
        writer.WriteEndElement();
    }

    private static void WriteRootRelationships(PoiXmlWriter writer)
    {
        writer.WriteStartElement("Relationships");
        writer.WriteAttributeString("xmlns", "http://schemas.openxmlformats.org/package/2006/relationships");
        WriteRelationship(writer, "rId1", "xl/workbook.xml", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument");
        WriteRelationship(writer, "rId2", "docProps/app.xml", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties");
        WriteRelationship(writer, "rId3", "docProps/core.xml", "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties");
        writer.WriteEndElement();
    }

    private void WriteSheetRelationships(PoiXmlWriter writer, XSSFSheet sheet)
    {
        writer.WriteStartElement("Relationships");
        writer.WriteAttributeString("xmlns", "http://schemas.openxmlformats.org/package/2006/relationships");
        int relId = 1;
        if (sheet.Drawing is not null)
        {
            WriteRelationship(writer, $"rId{relId++}", $"../drawings/drawing{sheet.Drawing!.DrawingIndex}.xml", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing");
        }
        foreach (var hyperlink in sheet.Hyperlinks)
        {
            var targetMode = hyperlink.IsExternal ? "External" : null;
            WriteRelationship(writer, hyperlink.RelationshipId ?? $"rId{relId++}", hyperlink.Address, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink", targetMode);
        }
        // Pivot table relationships
        foreach (var pt in sheet.PivotTables)
        {
            WriteRelationship(writer, $"rId{relId++}", $"../pivotTables/pivotTable{pt.PivotTableIndex}.xml",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable");
        }
        writer.WriteEndElement();
    }

    private void WriteDrawingRelationships(PoiXmlWriter writer, XSSFDrawing drawing)
    {
        writer.WriteStartElement("Relationships");
        writer.WriteAttributeString("xmlns", "http://schemas.openxmlformats.org/package/2006/relationships");
        foreach (var picture in drawing.Pictures)
        {
            var pictureData = GetPictureData(picture.PictureIndex);
            WriteRelationship(writer, picture.RelationshipId, $"../media/image{pictureData.Index}.{pictureData.Extension}", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image");
        }
        writer.WriteEndElement();
    }

    private void WriteDrawing(PoiXmlWriter writer, XSSFDrawing drawing)
    {
        writer.WriteStartElement("xdr", "wsDr");
        writer.WriteAttributeString("xmlns:xdr", "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing");
        writer.WriteAttributeString("xmlns:a", "http://schemas.openxmlformats.org/drawingml/2006/main");
        writer.WriteAttributeString("xmlns:r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
        foreach (var picture in drawing.Pictures)
        {
            WritePictureAnchor(writer, picture);
        }
        writer.WriteEndElement();
    }

    private static void WritePictureAnchor(PoiXmlWriter writer, XSSFPicture picture)
    {
        var anchor = picture.Anchor;
        writer.WriteStartElement("xdr", "twoCellAnchor");
        writer.WriteAttributeString("editAs", GetEditAs(anchor.getAnchorType()));
        WriteMarker(writer, "from", anchor.Col1, anchor.Dx1, anchor.Row1, anchor.Dy1);
        WriteMarker(writer, "to", anchor.Col2, anchor.Dx2, anchor.Row2, anchor.Dy2);
        writer.WriteStartElement("xdr", "pic");
        writer.WriteStartElement("xdr", "nvPicPr");
        writer.WriteStartElement("xdr", "cNvPr");
        writer.WriteAttributeString("id", picture.ShapeId.ToString(CultureInfo.InvariantCulture));
        writer.WriteAttributeString("name", "Picture " + picture.ShapeId.ToString(CultureInfo.InvariantCulture));
        writer.WriteAttributeString("descr", "Picture");
        writer.WriteEndElement();
        writer.WriteStartElement("xdr", "cNvPicPr");
        writer.WriteStartElement("a", "picLocks");
        writer.WriteAttributeString("noChangeAspect", "true");
        writer.WriteEndElement();
        writer.WriteEndElement();
        writer.WriteEndElement();
        writer.WriteStartElement("xdr", "blipFill");
        writer.WriteStartElement("a", "blip");
        writer.WriteAttributeString("r", "embed", picture.RelationshipId);
        writer.WriteEndElement();
        writer.WriteStartElement("a", "stretch");
        writer.WriteStartElement("a", "fillRect");
        writer.WriteEndElement();
        writer.WriteEndElement();
        writer.WriteEndElement();
        writer.WriteStartElement("xdr", "spPr");
        writer.WriteStartElement("a", "xfrm");
        if (picture.RotationAttribute != 0)
        {
            writer.WriteAttributeString("rot", picture.RotationAttribute.ToString(CultureInfo.InvariantCulture));
        }
        writer.WriteStartElement("a", "off");
        writer.WriteAttributeString("x", anchor.Dx1.ToString(CultureInfo.InvariantCulture));
        writer.WriteAttributeString("y", anchor.Dy1.ToString(CultureInfo.InvariantCulture));
        writer.WriteEndElement();
        writer.WriteStartElement("a", "ext");
        writer.WriteAttributeString("cx", Math.Max(0, anchor.Dx2 - anchor.Dx1).ToString(CultureInfo.InvariantCulture));
        writer.WriteAttributeString("cy", Math.Max(0, anchor.Dy2 - anchor.Dy1).ToString(CultureInfo.InvariantCulture));
        writer.WriteEndElement();
        writer.WriteEndElement();
        writer.WriteStartElement("a", "prstGeom");
        writer.WriteAttributeString("prst", "rect");
        writer.WriteStartElement("a", "avLst");
        writer.WriteEndElement();
        writer.WriteEndElement();
        writer.WriteEndElement();
        writer.WriteEndElement();
        writer.WriteStartElement("xdr", "clientData");
        writer.WriteEndElement();
        writer.WriteEndElement();
    }

    private static void WriteMarker(PoiXmlWriter writer, string markerName, int column, int columnOffset, int row, int rowOffset)
    {
        writer.WriteStartElement("xdr", markerName);
        WriteTextElement(writer, "xdr", "col", column.ToString(CultureInfo.InvariantCulture));
        WriteTextElement(writer, "xdr", "colOff", columnOffset.ToString(CultureInfo.InvariantCulture));
        WriteTextElement(writer, "xdr", "row", row.ToString(CultureInfo.InvariantCulture));
        WriteTextElement(writer, "xdr", "rowOff", rowOffset.ToString(CultureInfo.InvariantCulture));
        writer.WriteEndElement();
    }

    private static void WriteTextElement(PoiXmlWriter writer, string prefix, string localName, string text)
    {
        writer.WriteStartElement(prefix, localName);
        writer.WriteString(text);
        writer.WriteEndElement();
    }

    private static void WriteAppProperties(PoiXmlWriter writer)
    {
        writer.WriteStartElement("Properties");
        writer.WriteAttributeString("xmlns", "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties");
        writer.WriteStartElement("Application");
        writer.WriteString("Apache POI");
        writer.WriteEndElement();
        writer.WriteEndElement();
    }

    private static void WriteFont(PoiXmlWriter writer, XSSFFont font)
    {
        writer.WriteStartElement("font");
        if (font.getBold())
        {
            writer.WriteStartElement("b");
            writer.WriteEndElement();
        }

        if (font.getItalic())
        {
            writer.WriteStartElement("i");
            writer.WriteEndElement();
        }

        if (font.getStrikeout())
        {
            writer.WriteStartElement("strike");
            writer.WriteEndElement();
        }

        if (font.getUnderline() != 0)
        {
            writer.WriteStartElement("u");
            writer.WriteEndElement();
        }

        WriteValElement(writer, "sz", font.GetFontHeightText());
        writer.WriteStartElement("color");
        writer.WriteAttributeString("indexed", font.getColor().ToString(CultureInfo.InvariantCulture));
        writer.WriteEndElement();
        WriteValElement(writer, "name", font.getFontName());
        WriteValElement(writer, "family", "2");
        WriteValElement(writer, "scheme", "minor");
        writer.WriteEndElement();
    }

    private static void WriteFill(PoiXmlWriter writer, XSSFCellStyle? style)
    {
        writer.WriteStartElement("fill");
        writer.WriteStartElement("patternFill");
        writer.WriteAttributeString("patternType", GetPatternName(style?.FillPattern ?? FillPatternType.NoFill));
        if (style?.FillForegroundColor is short foreground)
        {
            writer.WriteStartElement("fgColor");
            writer.WriteAttributeString("indexed", foreground.ToString(CultureInfo.InvariantCulture));
            writer.WriteEndElement();
        }
        writer.WriteEndElement();
        writer.WriteEndElement();
    }

    private static void WriteBorder(PoiXmlWriter writer, XSSFCellStyle? style)
    {
        writer.WriteStartElement("border");
        WriteBorderSide(writer, "left", style?.BorderLeft ?? BorderStyle.None);
        WriteBorderSide(writer, "right", style?.BorderRight ?? BorderStyle.None);
        WriteBorderSide(writer, "top", style?.BorderTop ?? BorderStyle.None);
        WriteBorderSide(writer, "bottom", style?.BorderBottom ?? BorderStyle.None);
        writer.WriteStartElement("diagonal");
        writer.WriteEndElement();
        writer.WriteEndElement();
    }

    private static void WriteBorderSide(PoiXmlWriter writer, string sideName, BorderStyle style)
    {
        writer.WriteStartElement(sideName);
        if (style != BorderStyle.None)
        {
            writer.WriteAttributeString("style", GetBorderStyleName(style));
        }
        writer.WriteEndElement();
    }

    private static void WriteCellXf(PoiXmlWriter writer, XSSFCellStyle style)
    {
        writer.WriteStartElement("xf");
        writer.WriteAttributeString("numFmtId", style.NumFmtId.ToString(CultureInfo.InvariantCulture));
        writer.WriteAttributeString("fontId", style.FontId.ToString(CultureInfo.InvariantCulture));
        writer.WriteAttributeString("fillId", style.FillId.ToString(CultureInfo.InvariantCulture));
        writer.WriteAttributeString("borderId", style.BorderId.ToString(CultureInfo.InvariantCulture));
        writer.WriteAttributeString("xfId", "0");
        if (style.ApplyFont)
        {
            writer.WriteAttributeString("applyFont", "true");
        }
        if (style.ApplyNumberFormat)
        {
            writer.WriteAttributeString("applyNumberFormat", "true");
        }
        if (style.ApplyFill)
        {
            writer.WriteAttributeString("applyFill", "true");
        }
        if (style.ApplyBorder)
        {
            writer.WriteAttributeString("applyBorder", "true");
        }
        if (style.ApplyAlignment)
        {
            writer.WriteAttributeString("applyAlignment", "true");
        }
        if (style.ApplyAlignment)
        {
            writer.WriteStartElement("alignment");
            if (style.AlignmentValue != HorizontalAlignment.General)
            {
                writer.WriteAttributeString("horizontal", GetHorizontalAlignmentName(style.AlignmentValue));
            }
            if (style.VerticalAlignmentValue != VerticalAlignment.Bottom)
            {
                writer.WriteAttributeString("vertical", GetVerticalAlignmentName(style.VerticalAlignmentValue));
            }
            if (style.WrapTextEnabled)
            {
                writer.WriteAttributeString("wrapText", "true");
            }
            if (style.IndentLevel != 0)
            {
                writer.WriteAttributeString("indent", style.IndentLevel.ToString(CultureInfo.InvariantCulture));
            }
            if (style.TextRotation != 0)
            {
                writer.WriteAttributeString("textRotation", style.TextRotation.ToString(CultureInfo.InvariantCulture));
            }
            writer.WriteEndElement();
        }
        writer.WriteEndElement();
    }

    private static void WriteCoreProperties(PoiXmlWriter writer)
    {
        writer.WriteStartElement("cp", "coreProperties");
        writer.WriteAttributeString("xmlns:cp", "http://schemas.openxmlformats.org/package/2006/metadata/core-properties");
        writer.WriteAttributeString("xmlns:dc", "http://purl.org/dc/elements/1.1/");
        writer.WriteAttributeString("xmlns:dcterms", "http://purl.org/dc/terms/");
        writer.WriteAttributeString("xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance");
        writer.WriteStartElement("dcterms", "created");
        writer.WriteAttributeString("xsi", "type", "dcterms:W3CDTF");
        writer.WriteString(DateTime.UtcNow.ToString("yyyy-MM-dd'T'HH:mm:ss'Z'", CultureInfo.InvariantCulture));
        writer.WriteEndElement();
        writer.WriteStartElement("dc", "creator");
        writer.WriteString("Apache POI");
        writer.WriteEndElement();
        writer.WriteEndElement();
    }

    private void WriteWorkbook(PoiXmlWriter writer)
    {
        writer.WriteStartElement("workbook");
        writer.WriteAttributeString("xmlns", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
        writer.WriteAttributeString("xmlns:r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
        writer.WriteStartElement("workbookPr");
        writer.WriteAttributeString("date1904", "false");
        writer.WriteEndElement();
        if (_workbookProtected)
        {
            writer.WriteStartElement("workbookProtection");
            writer.WriteAttributeString("lockStructure", "1");
            writer.WriteEndElement();
        }
        writer.WriteStartElement("bookViews");
        writer.WriteStartElement("workbookView");
        writer.WriteAttributeString("activeTab", "0");
        writer.WriteEndElement();
        writer.WriteEndElement();
        writer.WriteStartElement("sheets");
        foreach (var sheet in _sheets)
        {
            writer.WriteStartElement("sheet");
            writer.WriteAttributeString("name", sheet.SheetName);
            writer.WriteAttributeString("r", "id", "rId" + (sheet.SheetIndex + 2).ToString(CultureInfo.InvariantCulture));
            writer.WriteAttributeString("sheetId", sheet.SheetIndex.ToString(CultureInfo.InvariantCulture));
            writer.WriteEndElement();
        }
        writer.WriteEndElement(); // sheets

        // pivotCaches
        if (_pivotTables.Count > 0)
        {
            writer.WriteStartElement("pivotCaches");
            foreach (var pt in _pivotTables)
            {
                writer.WriteStartElement("pivotCache");
                writer.WriteAttributeString("cacheId", pt.CacheId.ToString(CultureInfo.InvariantCulture));
                writer.WriteAttributeString("r", "id", pt.CacheRelId ?? $"rId{_sheets.Count + 2 + pt.PivotTableIndex}");
                writer.WriteEndElement();
            }
            writer.WriteEndElement();
        }

        if (_hasCalcPr)
        {
            writer.WriteStartElement("calcPr");
            writer.WriteAttributeString("calcId", "0");
            writer.WriteAttributeString("fullCalcOnLoad", _forceFormulaRecalculation ? "true" : "false");
            writer.WriteEndElement();
        }
        writer.WriteEndElement();
    }

    private void WriteWorkbookRelationships(PoiXmlWriter writer)
    {
        writer.WriteStartElement("Relationships");
        writer.WriteAttributeString("xmlns", "http://schemas.openxmlformats.org/package/2006/relationships");
        WriteRelationship(writer, "rId1", "sharedStrings.xml", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings");
        WriteRelationship(writer, "rId2", "styles.xml", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles");
        int nextRelId = 3;
        foreach (var sheet in _sheets)
        {
            WriteRelationship(writer, "rId" + nextRelId.ToString(CultureInfo.InvariantCulture), $"worksheets/sheet{sheet.SheetIndex}.xml", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet");
            nextRelId++;
        }
        // xlsm: append vbaProject relationship after all worksheet rels
        if (_vbaProjectBin != null)
        {
            WriteRelationship(writer, "rId" + nextRelId.ToString(CultureInfo.InvariantCulture), "vbaProject.bin", RelTypeVbaProject);
            nextRelId++;
        }
        // Pivot cache definition relationships
        foreach (var pt in _pivotTables)
        {
            var relId = "rId" + nextRelId.ToString(CultureInfo.InvariantCulture);
            pt.CacheRelId = relId;
            nextRelId++;
            WriteRelationship(writer, relId, $"pivotCache/pivotCacheDefinition{pt.CacheId + 1}.xml",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition");
        }
        writer.WriteEndElement();
    }

    private void WriteStyles(PoiXmlWriter writer)
    {
        writer.WriteStartElement("styleSheet");
        writer.WriteAttributeString("xmlns", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
        writer.WriteStartElement("numFmts");
        writer.WriteAttributeString("count", _customNumberFormats.Count.ToString(CultureInfo.InvariantCulture));
        foreach (var format in _customNumberFormats)
        {
            writer.WriteStartElement("numFmt");
            writer.WriteAttributeString("numFmtId", format.Key.ToString(CultureInfo.InvariantCulture));
            writer.WriteAttributeString("formatCode", format.Value);
            writer.WriteEndElement();
        }
        writer.WriteEndElement();
        writer.WriteStartElement("fonts");
        writer.WriteAttributeString("count", _fonts.Count.ToString(CultureInfo.InvariantCulture));
        foreach (var font in _fonts)
        {
            WriteFont(writer, font);
        }
        writer.WriteEndElement();
        writer.WriteStartElement("fills");
        writer.WriteAttributeString("count", _fills.Count.ToString(CultureInfo.InvariantCulture));
        writer.WriteStartElement("fill");
        writer.WriteStartElement("patternFill");
        writer.WriteAttributeString("patternType", "none");
        writer.WriteEndElement();
        writer.WriteEndElement();
        writer.WriteStartElement("fill");
        writer.WriteStartElement("patternFill");
        writer.WriteAttributeString("patternType", "darkGray");
        writer.WriteEndElement();
        writer.WriteEndElement();
        for (var i = 2; i < _fills.Count; i++)
        {
            WriteFill(writer, _fills[i]);
        }
        writer.WriteEndElement();
        writer.WriteStartElement("borders");
        writer.WriteAttributeString("count", _borders.Count.ToString(CultureInfo.InvariantCulture));
        WriteBorder(writer, null);
        for (var i = 1; i < _borders.Count; i++)
        {
            WriteBorder(writer, _borders[i]);
        }
        writer.WriteEndElement();
        writer.WriteStartElement("cellStyleXfs");
        writer.WriteAttributeString("count", "1");
        writer.WriteStartElement("xf");
        writer.WriteAttributeString("numFmtId", "0");
        writer.WriteAttributeString("fontId", "0");
        writer.WriteAttributeString("fillId", "0");
        writer.WriteAttributeString("borderId", "0");
        writer.WriteEndElement();
        writer.WriteEndElement();
        writer.WriteStartElement("cellXfs");
        writer.WriteAttributeString("count", _cellStyles.Count.ToString(CultureInfo.InvariantCulture));
        foreach (var style in _cellStyles)
        {
            WriteCellXf(writer, style);
        }
        writer.WriteEndElement();
        writer.WriteEndElement();
    }

    private void WriteSharedStrings(PoiXmlWriter writer)
    {
        writer.WriteStartElement("sst");
        writer.WriteAttributeString("count", CountStringCells().ToString(CultureInfo.InvariantCulture));
        writer.WriteAttributeString("uniqueCount", _sharedStrings.Count.ToString(CultureInfo.InvariantCulture));
        writer.WriteAttributeString("xmlns", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
        foreach (var rts in _sharedStrings)
        {
            writer.WriteStartElement("si");

            if (rts.IsRichText)
            {
                // Write si with <r> elements for formatted runs
                foreach (var run in rts.Runs)
                {
                    writer.WriteStartElement("r");

                    // Write rPr if any formatting is set
                    if (run.Bold || run.Italic || run.Underline || run.Strikethrough
                        || run.FontSize > 0 || run.FontName is not null || run.Color is not null)
                    {
                        writer.WriteStartElement("rPr");

                        if (run.Bold) { writer.WriteStartElement("b"); writer.WriteEndElement(); }
                        if (run.Italic) { writer.WriteStartElement("i"); writer.WriteEndElement(); }
                        if (run.Underline) { writer.WriteStartElement("u"); writer.WriteEndElement(); }
                        if (run.Strikethrough) { writer.WriteStartElement("strike"); writer.WriteEndElement(); }

                        if (run.FontSize > 0)
                        {
                            writer.WriteStartElement("sz");
                            writer.WriteAttributeString("val", ((int)(run.FontSize * 100)).ToString(CultureInfo.InvariantCulture));
                            writer.WriteEndElement();
                        }

                        if (run.FontName is not null)
                        {
                            writer.WriteStartElement("rFont");
                            writer.WriteAttributeString("val", run.FontName);
                            writer.WriteEndElement();
                        }

                        if (run.Color is not null)
                        {
                            writer.WriteStartElement("color");
                            writer.WriteAttributeString("rgb", run.Color);
                            writer.WriteEndElement();
                        }

                        writer.WriteEndElement(); // rPr
                    }

                    writer.WriteStartElement("t");
                    writer.WriteString(run.Text);
                    writer.WriteEndElement(); // t
                    writer.WriteEndElement(); // r
                }
            }
            else
            {
                // Plain text: single <t> element
                writer.WriteStartElement("t");
                writer.WriteString(rts.getString());
                writer.WriteEndElement();
            }

            writer.WriteEndElement(); // si
        }
        writer.WriteEndElement(); // sst
    }

    private int CountStringCells()
    {
        var count = 0;
        foreach (var sheet in _sheets)
        {
            foreach (var row in sheet.Rows)
            {
                count += row.Cells.Count(cell => cell.getCellType() == CellType.String);
            }
        }

        return count;
    }

    private void WriteWorksheet(PoiXmlWriter writer, XSSFSheet sheet)
    {
        writer.WriteStartElement("worksheet");
        writer.WriteAttributeString("xmlns", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
        writer.WriteAttributeString("xmlns:r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
        writer.WriteStartElement("dimension");
        writer.WriteAttributeString("ref", GetDimensionReference(sheet));
        writer.WriteEndElement();
        writer.WriteStartElement("sheetViews");
        writer.WriteStartElement("sheetView");
        writer.WriteAttributeString("tabSelected", "1");
        writer.WriteAttributeString("workbookViewId", "0");
        if (sheet.FreezeColSplit > 0 || sheet.FreezeRowSplit > 0)
        {
            writer.WriteStartElement("pane");
            if (sheet.FreezeColSplit > 0)
                writer.WriteAttributeString("xSplit", sheet.FreezeColSplit.ToString(CultureInfo.InvariantCulture));
            if (sheet.FreezeRowSplit > 0)
                writer.WriteAttributeString("ySplit", sheet.FreezeRowSplit.ToString(CultureInfo.InvariantCulture));
            writer.WriteAttributeString("state", "frozen");
            var activePane = (sheet.FreezeRowSplit > 0, sheet.FreezeColSplit > 0) switch
            {
                (true, false) => "bottomLeft",
                (false, true) => "topRight",
                _             => "bottomRight"
            };
            writer.WriteAttributeString("activePane", activePane);
            {
                var topRow = sheet.FreezeRowSplit > 0 ? sheet.FreezeRowSplit : 0;
                var leftCol = sheet.FreezeColSplit > 0 ? sheet.FreezeColSplit : 0;
                writer.WriteAttributeString("topLeftCell", FormatCellReference(topRow, leftCol));
            }
            writer.WriteEndElement();
        }
        writer.WriteEndElement();
        writer.WriteEndElement();
        writer.WriteStartElement("sheetFormatPr");
        writer.WriteAttributeString("defaultRowHeight", "15.0");
        writer.WriteEndElement();
        if (sheet.SheetProtected)
        {
            writer.WriteStartElement("sheetProtection");
            writer.WriteAttributeString("sheet", "1");
            writer.WriteAttributeString("objects", "1");
            writer.WriteAttributeString("scenarios", "1");
            writer.WriteEndElement();
        }
        WriteCols(writer, sheet);
        writer.WriteStartElement("sheetData");
        foreach (var row in sheet.Rows)
        {
            WriteRow(writer, row);
        }
        writer.WriteEndElement();
        WriteMergeCells(writer, sheet);
        WriteHyperlinks(writer, sheet);
        WriteDataValidations(writer, sheet);
        WriteConditionalFormatting(writer, sheet);
        if (sheet.AutoFilter is not null)
        {
            writer.WriteStartElement("autoFilter");
            writer.WriteAttributeString("ref", sheet.AutoFilter.FormatAsString());
            writer.WriteEndElement();
        }
        writer.WriteStartElement("pageMargins");
        writer.WriteAttributeString("bottom", sheet.PageMarginBottom.ToString("F4", CultureInfo.InvariantCulture));
        writer.WriteAttributeString("footer", sheet.PageMarginFooter.ToString("F4", CultureInfo.InvariantCulture));
        writer.WriteAttributeString("header", sheet.PageMarginHeader.ToString("F4", CultureInfo.InvariantCulture));
        writer.WriteAttributeString("left", sheet.PageMarginLeft.ToString("F4", CultureInfo.InvariantCulture));
        writer.WriteAttributeString("right", sheet.PageMarginRight.ToString("F4", CultureInfo.InvariantCulture));
        writer.WriteAttributeString("top", sheet.PageMarginTop.ToString("F4", CultureInfo.InvariantCulture));
        writer.WriteEndElement();
        if (sheet.PageOrientation is not null || sheet.FitToWidth is not null || sheet.FitToHeight is not null)
        {
            writer.WriteStartElement("pageSetup");
            if (sheet.PageOrientation is not null)
                writer.WriteAttributeString("orientation", sheet.PageOrientation);
            if (sheet.FitToWidth is not null)
                writer.WriteAttributeString("fitToWidth", sheet.FitToWidth.Value.ToString(CultureInfo.InvariantCulture));
            if (sheet.FitToHeight is not null)
                writer.WriteAttributeString("fitToHeight", sheet.FitToHeight.Value.ToString(CultureInfo.InvariantCulture));
            if (sheet.PaperSize is not null)
                writer.WriteAttributeString("paperSize", sheet.PaperSize.Value.ToString(CultureInfo.InvariantCulture));
            writer.WriteEndElement();
        }
        if (sheet.HeaderCenter is not null || sheet.FooterCenter is not null)
        {
            writer.WriteStartElement("headerFooter");
            if (sheet.HeaderCenter is not null)
            {
                writer.WriteStartElement("oddHeader");
                writer.WriteString(sheet.HeaderCenter);
                writer.WriteEndElement();
            }
            if (sheet.FooterCenter is not null)
            {
                writer.WriteStartElement("oddFooter");
                writer.WriteString(sheet.FooterCenter);
                writer.WriteEndElement();
            }
            writer.WriteEndElement();
        }
        if (sheet.Drawing is not null)
        {
            writer.WriteStartElement("drawing");
            writer.WriteAttributeString("r", "id", "rId1");
            writer.WriteEndElement();
        }
        writer.WriteEndElement();
    }

    private void WriteRow(PoiXmlWriter writer, XSSFRow row)
    {
        writer.WriteStartElement("row");
        writer.WriteAttributeString("r", (row.getRowNum() + 1).ToString(CultureInfo.InvariantCulture));
        if (row.HasCustomHeight)
            writer.WriteAttributeString("ht", row.HeightValue.ToString("F1", CultureInfo.InvariantCulture));
        if (row.HasCustomHeight)
            writer.WriteAttributeString("customHeight", "true");
        if (row.IsHidden)
            writer.WriteAttributeString("hidden", "1");
        foreach (var cell in row.Cells)
        {
            var isFormula = cell.getCellType() == CellType.Formula;
            var effectiveType = isFormula
                ? cell.getCachedFormulaResultType()
                : cell.getCellType();

            if (!isFormula && effectiveType == CellType.Blank) continue;

            // POI attribute order: r, t (if not numeric default), s (if non-zero)
            // Ported from XSSFCell.write() via XMLBeans schema order.
            writer.WriteStartElement("c");
            writer.WriteAttributeString("r", FormatCellReference(cell.getRowIndex(), cell.getColumnIndex()));

            // Write t before s — matches POI/XMLBeans output order
            var tAttr = isFormula && effectiveType == CellType.String
                ? "str"
                : effectiveType switch
            {
                CellType.String  => "s",
                CellType.Boolean => "b",
                CellType.Error   => "e",
                _                => null   // Numeric: t="n" is the OOXML default, omit it
            };
            if (tAttr is not null)
                writer.WriteAttributeString("t", tAttr);

            var styleIndex = cell.getCellStyle().getIndex();
            if (styleIndex != 0)
                writer.WriteAttributeString("s", styleIndex.ToString(CultureInfo.InvariantCulture));

            if (isFormula)
            {
                writer.WriteStartElement("f");
                writer.WriteString(cell.getCellFormula() ?? string.Empty);
                writer.WriteEndElement();
            }

            if (!isFormula || cell.HasCachedValue)
            {
                switch (effectiveType)
                {
                    case CellType.String:
                        writer.WriteStartElement("v");
                        writer.WriteString(isFormula
                            ? cell.getStringCellValue()
                            : GetSharedStringIndex(cell.getStringCellValue()).ToString(CultureInfo.InvariantCulture));
                        writer.WriteEndElement();
                        break;

                    case CellType.Boolean:
                        writer.WriteStartElement("v");
                        writer.WriteString(cell.getBooleanCellValue() ? "1" : "0");
                        writer.WriteEndElement();
                        break;

                    case CellType.Error:
                        writer.WriteStartElement("v");
                        writer.WriteString(cell.getErrorCellString());
                        writer.WriteEndElement();
                        break;

                    default: // Numeric
                        writer.WriteStartElement("v");
                        writer.WriteString(cell.GetNumericText());
                        writer.WriteEndElement();
                        break;
                }
            }

            writer.WriteEndElement(); // </c>
        }
        writer.WriteEndElement();
    }

    private static void WriteMergeCells(PoiXmlWriter writer, XSSFSheet sheet)
    {
        var regions = sheet.MergedRegions;
        if (regions.Count == 0) return;
        writer.WriteStartElement("mergeCells");
        writer.WriteAttributeString("count", regions.Count.ToString(CultureInfo.InvariantCulture));
        foreach (var region in regions)
        {
            writer.WriteStartElement("mergeCell");
            writer.WriteAttributeString("ref", region.FormatAsString());
            writer.WriteEndElement();
        }
        writer.WriteEndElement();
    }

    private static void WriteHyperlinks(PoiXmlWriter writer, XSSFSheet sheet)
    {
        var hyperlinks = sheet.Hyperlinks;
        if (hyperlinks.Count == 0) return;
        writer.WriteStartElement("hyperlinks");
        foreach (var hyperlink in hyperlinks)
        {
            writer.WriteStartElement("hyperlink");
            writer.WriteAttributeString("ref", hyperlink.CellRef);
            writer.WriteAttributeString("r", "id", hyperlink.RelationshipId ?? string.Empty);
            writer.WriteEndElement();
        }
        writer.WriteEndElement();
    }

    private static void WriteDataValidations(PoiXmlWriter writer, XSSFSheet sheet)
    {
        var validations = sheet.DataValidations;
        if (validations.Count == 0) return;
        writer.WriteStartElement("dataValidations");
        writer.WriteAttributeString("count", validations.Count.ToString(CultureInfo.InvariantCulture));
        foreach (var dv in validations)
        {
            writer.WriteStartElement("dataValidation");
            writer.WriteAttributeString("type", GetDataValidationTypeName(dv.Type));
            if (dv.Type != DataValidationType.List && dv.Type != DataValidationType.None)
                writer.WriteAttributeString("operator", GetDataValidationOperatorName(dv.Operator));
            if (!dv.AllowBlank) writer.WriteAttributeString("allowBlank", "0");
            if (!dv.ShowInputMessage) writer.WriteAttributeString("showInputMessage", "0");
            if (!dv.ShowErrorMessage) writer.WriteAttributeString("showErrorMessage", "0");
            if (!dv.ShowDropDown) writer.WriteAttributeString("showDropDown", "0");
            if (dv.ErrorStyle is not null) writer.WriteAttributeString("errorStyle", dv.ErrorStyle);
            if (dv.ErrorTitle is not null) writer.WriteAttributeString("errorTitle", dv.ErrorTitle);
            if (dv.ErrorMessage is not null) writer.WriteAttributeString("error", dv.ErrorMessage);
            if (dv.PromptTitle is not null) writer.WriteAttributeString("promptTitle", dv.PromptTitle);
            if (dv.PromptMessage is not null) writer.WriteAttributeString("prompt", dv.PromptMessage);
            writer.WriteAttributeString("sqref", dv.Sqref);
            if (dv.Formula1 is not null) { writer.WriteStartElement("formula1"); writer.WriteString(dv.Formula1); writer.WriteEndElement(); }
            if (dv.Formula2 is not null) { writer.WriteStartElement("formula2"); writer.WriteString(dv.Formula2); writer.WriteEndElement(); }
            writer.WriteEndElement();
        }
        writer.WriteEndElement();
    }

    private static void WriteConditionalFormatting(PoiXmlWriter writer, XSSFSheet sheet)
    {
        var cfs = sheet.ConditionalFormatting;
        if (cfs.Count == 0) return;
        foreach (var cf in cfs)
        {
            writer.WriteStartElement("conditionalFormatting");
            writer.WriteAttributeString("sqref", cf.Sqref);
            foreach (var rule in cf.Rules)
            {
                writer.WriteStartElement("cfRule");
                writer.WriteAttributeString("type", GetCfTypeName(rule.Type));
                writer.WriteAttributeString("priority", rule.Priority.ToString(CultureInfo.InvariantCulture));
                if (rule.Operator is not null)
                    writer.WriteAttributeString("operator", rule.Operator);
                if (rule.Text is not null)
                    writer.WriteAttributeString("text", rule.Text);
                if (rule.DxfId >= 0)
                    writer.WriteAttributeString("dxfId", rule.DxfId.ToString(CultureInfo.InvariantCulture));
                foreach (var formula in rule.Formulas)
                {
                    writer.WriteStartElement("formula");
                    writer.WriteString(formula);
                    writer.WriteEndElement();
                }
                writer.WriteEndElement();
            }
            writer.WriteEndElement();
        }
    }

    private static void WriteCols(PoiXmlWriter writer, XSSFSheet sheet)
    {
        var widths = sheet.ColumnWidths;
        var hiddenCols = sheet.HiddenColumns;
        var allCols = new HashSet<int>(widths.Keys);
        allCols.UnionWith(hiddenCols);
        if (allCols.Count == 0) return;
        writer.WriteStartElement("cols");
        foreach (var colIndex in allCols.OrderBy(c => c))
        {
            widths.TryGetValue(colIndex, out var width);
            var isHidden = hiddenCols.Contains(colIndex);
            writer.WriteStartElement("col");
            writer.WriteAttributeString("min", (colIndex + 1).ToString(CultureInfo.InvariantCulture));
            writer.WriteAttributeString("max", (colIndex + 1).ToString(CultureInfo.InvariantCulture));
            if (width > 0)
            {
                writer.WriteAttributeString("width", (width / 256.0).ToString("F2", CultureInfo.InvariantCulture));
                writer.WriteAttributeString("customWidth", "true");
            }
            if (isHidden)
                writer.WriteAttributeString("hidden", "1");
            writer.WriteEndElement();
        }
        writer.WriteEndElement();
    }

    private static string GetDimensionReference(XSSFSheet sheet)
    {
        if (sheet.Rows.Count == 0)
        {
            return "A1";
        }

        var minRow = int.MaxValue;
        var maxRow = 0;
        var minColumn = int.MaxValue;
        var maxColumn = 0;

        foreach (var row in sheet.Rows)
        {
            minRow = Math.Min(minRow, row.getRowNum());
            maxRow = Math.Max(maxRow, row.getRowNum());
            foreach (var cell in row.Cells)
            {
                var dimType = cell.getCellType() == CellType.Formula
                    ? cell.getCachedFormulaResultType() : cell.getCellType();
                if (dimType == CellType.Blank) continue;

                minColumn = Math.Min(minColumn, cell.getColumnIndex());
                maxColumn = Math.Max(maxColumn, cell.getColumnIndex());
            }
        }

        if (minColumn == int.MaxValue)
        {
            return "A1";
        }

        var first = FormatCellReference(minRow, minColumn);
        var last = FormatCellReference(maxRow, maxColumn);
        return first == last ? first : string.Concat(first, ":", last);
    }

    private static string FormatCellReference(int rowIndex, int columnIndex)
    {
        return string.Concat(FormatColumnName(columnIndex), (rowIndex + 1).ToString(CultureInfo.InvariantCulture));
    }

    private static string FormatColumnName(int columnIndex)
    {
        var column = columnIndex + 1;
        var name = string.Empty;
        while (column > 0)
        {
            column--;
            name = (char)('A' + column % 26) + name;
            column /= 26;
        }

        return name;
    }

    private static int ParseColumnIndex(string cellReference)
    {
        var columnIndex = 0;
        var seenColumn = false;

        foreach (var character in cellReference)
        {
            if (!char.IsLetter(character))
            {
                break;
            }

            seenColumn = true;
            columnIndex = (columnIndex * 26) + char.ToUpperInvariant(character) - 'A' + 1;
        }

        if (!seenColumn)
        {
            throw new FormatException($"Cell reference '{cellReference}' does not contain a column.");
        }

        return columnIndex - 1;
    }

    private static bool ParseBooleanAttribute(string? value, bool defaultValue)
    {
        if (value is null)
        {
            return defaultValue;
        }

        return value == "1" || string.Equals(value, "true", StringComparison.OrdinalIgnoreCase);
    }

    private static string GetPatternName(FillPatternType pattern)
    {
        return pattern switch
        {
            FillPatternType.NoFill => "none",
            FillPatternType.SolidForeground => "solid",
            FillPatternType.FineDots => "mediumGray",
            FillPatternType.AltBars => "darkGray",
            FillPatternType.SparseDots => "lightGray",
            FillPatternType.ThickHorizontalBands => "darkHorizontal",
            FillPatternType.ThickVerticalBands => "darkVertical",
            FillPatternType.ThickBackwardDiagonals => "darkDown",
            FillPatternType.ThickForwardDiagonals => "darkUp",
            FillPatternType.BigSpots => "darkGrid",
            FillPatternType.Bricks => "darkTrellis",
            FillPatternType.ThinHorizontalBands => "lightHorizontal",
            FillPatternType.ThinVerticalBands => "lightVertical",
            FillPatternType.ThinBackwardDiagonals => "lightDown",
            FillPatternType.ThinForwardDiagonals => "lightUp",
            FillPatternType.Squares => "lightGrid",
            FillPatternType.Diamonds => "lightTrellis",
            FillPatternType.LessDots => "gray125",
            FillPatternType.LeastDots => "gray0625",
            _ => "none"
        };
    }

    private static string GetHorizontalAlignmentName(HorizontalAlignment alignment)
    {
        return alignment switch
        {
            HorizontalAlignment.General        => "general",
            HorizontalAlignment.Left           => "left",
            HorizontalAlignment.Center         => "center",
            HorizontalAlignment.Right          => "right",
            HorizontalAlignment.Fill           => "fill",
            HorizontalAlignment.Justify        => "justify",
            HorizontalAlignment.CenterSelection => "centerContinuous",
            HorizontalAlignment.Distributed    => "distributed",
            _                                  => "general"
        };
    }

    private static string GetVerticalAlignmentName(VerticalAlignment alignment)
    {
        return alignment switch
        {
            VerticalAlignment.Bottom       => "bottom",
            VerticalAlignment.Center       => "center",
            VerticalAlignment.Top          => "top",
            VerticalAlignment.Justify      => "justify",
            VerticalAlignment.Distributed  => "distributed",
            _                              => "bottom"
        };
    }

    private static string GetBorderStyleName(BorderStyle style)
    {
        return style switch
        {
            BorderStyle.Thin => "thin",
            BorderStyle.Medium => "medium",
            BorderStyle.Dashed => "dashed",
            BorderStyle.Dotted => "dotted",
            BorderStyle.Thick => "thick",
            BorderStyle.Double => "double",
            BorderStyle.Hair => "hair",
            BorderStyle.MediumDashed => "mediumDashed",
            BorderStyle.DashDot => "dashDot",
            BorderStyle.MediumDashDot => "mediumDashDot",
            BorderStyle.DashDotDot => "dashDotDot",
            BorderStyle.MediumDashDotDot => "mediumDashDotDot",
            BorderStyle.SlantedDashDot => "slantDashDot",
            _ => "none"
        };
    }

    private static string GetEditAs(AnchorType anchorType)
    {
        return anchorType switch
        {
            AnchorType.DONT_MOVE_AND_RESIZE => "absolute",
            AnchorType.MOVE_DONT_RESIZE => "oneCell",
            _ => "twoCell"
        };
    }

    private static void ValidatePictureType(int format)
    {
        _ = format switch
        {
            PICTURE_TYPE_JPEG or PICTURE_TYPE_PNG or PICTURE_TYPE_DIB or PICTURE_TYPE_GIF or PICTURE_TYPE_TIFF or PICTURE_TYPE_EPS or PICTURE_TYPE_BMP or PICTURE_TYPE_WPG or PICTURE_TYPE_EMF => true,
            _ => throw new ArgumentException($"Picture type {format} is not supported by the Phase 2.5 XSSF writer.", nameof(format))
        };
    }

    private static int GetPictureTypeFromExtension(string extension)
    {
        return extension.TrimStart('.').ToLowerInvariant() switch
        {
            "jpg" or "jpeg" => PICTURE_TYPE_JPEG,
            "png" => PICTURE_TYPE_PNG,
            "dib" => PICTURE_TYPE_DIB,
            "gif" => PICTURE_TYPE_GIF,
            "tif" or "tiff" => PICTURE_TYPE_TIFF,
            "emf" => PICTURE_TYPE_EMF,
            "eps" => PICTURE_TYPE_EPS,
            "bmp" => PICTURE_TYPE_BMP,
            "wpg" => PICTURE_TYPE_WPG,
            _ => throw new NotSupportedException($"Picture extension '{extension}' is not supported by the Phase 2.5 XSSF reader.")
        };
    }

    private static void WriteDefault(PoiXmlWriter writer, string extension, string contentType)
    {
        writer.WriteStartElement("Default");
        writer.WriteAttributeString("Extension", extension);
        writer.WriteAttributeString("ContentType", contentType);
        writer.WriteEndElement();
    }

    private static void WriteOverride(PoiXmlWriter writer, string partName, string contentType)
    {
        writer.WriteStartElement("Override");
        writer.WriteAttributeString("PartName", partName);
        writer.WriteAttributeString("ContentType", contentType);
        writer.WriteEndElement();
    }

    private static void WriteRelationship(PoiXmlWriter writer, string id, string target, string type)
    {
        writer.WriteStartElement("Relationship");
        writer.WriteAttributeString("Id", id);
        writer.WriteAttributeString("Target", target);
        writer.WriteAttributeString("Type", type);
        writer.WriteEndElement();
    }

    private static void WriteRelationship(PoiXmlWriter writer, string id, string target, string type, string? targetMode)
    {
        writer.WriteStartElement("Relationship");
        writer.WriteAttributeString("Id", id);
        writer.WriteAttributeString("Target", target);
        writer.WriteAttributeString("Type", type);
        if (targetMode is not null)
            writer.WriteAttributeString("TargetMode", targetMode);
        writer.WriteEndElement();
    }

    private static void WriteValElement(PoiXmlWriter writer, string elementName, string value)
    {
        writer.WriteStartElement(elementName);
        writer.WriteAttributeString("val", value);
        writer.WriteEndElement();
    }

    private static double ParseDoubleAttr(XmlReader reader, string attrName, double defaultValue)
    {
        var text = reader.GetAttribute(attrName);
        if (text is not null && double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out var value))
            return value;
        return defaultValue;
    }

    private static string GetDataValidationTypeName(DataValidationType type)
    {
        return type switch
        {
            DataValidationType.Whole => "whole",
            DataValidationType.Decimal => "decimal",
            DataValidationType.List => "list",
            DataValidationType.Date => "date",
            DataValidationType.Time => "time",
            DataValidationType.TextLength => "textLength",
            DataValidationType.Custom => "custom",
            _ => "none"
        };
    }

    private static string GetDataValidationOperatorName(DataValidationOperator op)
    {
        return op switch
        {
            DataValidationOperator.NotBetween => "notBetween",
            DataValidationOperator.Equal => "equal",
            DataValidationOperator.NotEqual => "notEqual",
            DataValidationOperator.LessThan => "lessThan",
            DataValidationOperator.LessThanOrEqual => "lessThanOrEqual",
            DataValidationOperator.GreaterThan => "greaterThan",
            DataValidationOperator.GreaterThanOrEqual => "greaterThanOrEqual",
            _ => "between"
        };
    }

    private static DataValidationType DataValidationTypeFromName(string name)
    {
        return name switch
        {
            "whole" => DataValidationType.Whole,
            "decimal" => DataValidationType.Decimal,
            "list" => DataValidationType.List,
            "date" => DataValidationType.Date,
            "time" => DataValidationType.Time,
            "textLength" => DataValidationType.TextLength,
            "custom" => DataValidationType.Custom,
            _ => DataValidationType.None
        };
    }

    private static DataValidationOperator DataValidationOperatorFromName(string name)
    {
        return name switch
        {
            "notBetween" => DataValidationOperator.NotBetween,
            "equal" => DataValidationOperator.Equal,
            "notEqual" => DataValidationOperator.NotEqual,
            "lessThan" => DataValidationOperator.LessThan,
            "lessThanOrEqual" => DataValidationOperator.LessThanOrEqual,
            "greaterThan" => DataValidationOperator.GreaterThan,
            "greaterThanOrEqual" => DataValidationOperator.GreaterThanOrEqual,
            _ => DataValidationOperator.Between
        };
    }

    private static string GetCfTypeName(ConditionalFormatType type)
    {
        return type switch
        {
            ConditionalFormatType.CellIs => "cellIs",
            ConditionalFormatType.Formula => "expression",
            ConditionalFormatType.Top10 => "top10",
            ConditionalFormatType.UniqueValues => "uniqueValues",
            ConditionalFormatType.DuplicateValues => "duplicateValues",
            ConditionalFormatType.ContainsText => "containsText",
            ConditionalFormatType.NotContainsText => "notContainsText",
            ConditionalFormatType.BeginsWith => "beginsWith",
            ConditionalFormatType.EndsWith => "endsWith",
            ConditionalFormatType.ContainsBlanks => "containsBlanks",
            ConditionalFormatType.NotContainsBlanks => "notContainsBlanks",
            ConditionalFormatType.ContainsErrors => "containsErrors",
            ConditionalFormatType.NotContainsErrors => "notContainsErrors",
            ConditionalFormatType.TimePeriod => "timePeriod",
            ConditionalFormatType.AboveAverage => "aboveAverage",
            _ => "cellIs"
        };
    }

    private static ConditionalFormatType CfTypeFromName(string name)
    {
        return name switch
        {
            "cellIs" => ConditionalFormatType.CellIs,
            "expression" => ConditionalFormatType.Formula,
            "top10" => ConditionalFormatType.Top10,
            "uniqueValues" => ConditionalFormatType.UniqueValues,
            "duplicateValues" => ConditionalFormatType.DuplicateValues,
            "containsText" => ConditionalFormatType.ContainsText,
            "notContainsText" => ConditionalFormatType.NotContainsText,
            "beginsWith" => ConditionalFormatType.BeginsWith,
            "endsWith" => ConditionalFormatType.EndsWith,
            "containsBlanks" => ConditionalFormatType.ContainsBlanks,
            "notContainsBlanks" => ConditionalFormatType.NotContainsBlanks,
            "containsErrors" => ConditionalFormatType.ContainsErrors,
            "notContainsErrors" => ConditionalFormatType.NotContainsErrors,
            "timePeriod" => ConditionalFormatType.TimePeriod,
            "aboveAverage" => ConditionalFormatType.AboveAverage,
            _ => ConditionalFormatType.CellIs
        };
    }

    private static void WritePivotCacheDefinition(PoiXmlWriter writer, XSSFPivotCacheDefinition cacheDef)
    {
        writer.WriteStartElement("pivotCacheDefinition", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
        writer.WriteAttributeString("xmlns:r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
        writer.WriteAttributeString("r", "id", "");
        writer.WriteAttributeString("refreshOnLoad", "1");
        writer.WriteAttributeString("updatedBy", "DotnetPoi");
        writer.WriteAttributeString("createdVersion", "3");
        writer.WriteAttributeString("minRefreshableVersion", "3");
        writer.WriteAttributeString("recordCount", "0");
        writer.WriteAttributeString("upgradeOnRefresh", "1");

        writer.WriteStartElement("cacheSource");
        writer.WriteAttributeString("type", "worksheet");
        writer.WriteStartElement("worksheetSource");
        writer.WriteAttributeString("ref", cacheDef.SourceRef ?? string.Empty);
        writer.WriteAttributeString("sheet", cacheDef.SourceSheetName ?? string.Empty);
        writer.WriteEndElement();
        writer.WriteEndElement();

        writer.WriteEndElement();
    }

    private static void WritePivotCacheRecords(PoiXmlWriter writer, XSSFPivotCacheRecords records)
    {
        writer.WriteStartElement("pivotCacheRecords", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
        writer.WriteAttributeString("count", "0");
        writer.WriteEndElement();
    }

    private sealed record SheetInfo(string Name, string PartName);
}
