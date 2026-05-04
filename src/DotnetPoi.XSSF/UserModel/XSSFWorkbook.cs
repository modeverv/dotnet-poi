using System.Globalization;
using System.IO.Compression;
using System.Text;
using System.Xml;
using DotnetPoi.POIFS.Crypt;
using DotnetPoi.SS.UserModel;
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
    private readonly List<string> _sharedStrings = new();
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
    private string _calcId = "0";
    private bool _isDirty;
    private bool _isLoading;
    private Dictionary<string, byte[]>? _originalPackageEntries;

    // Unparsed parts support for preserving fidelity (e.g. theme, calcChain)
    private readonly List<UnparsedPartInfo> _unparsedParts = new();
    private readonly Dictionary<string, byte[]> _unparsedPartContents = new(StringComparer.Ordinal);

    // Preserved root-level document properties (docProps/app.xml, docProps/core.xml)
    private byte[]? _originalDocPropsApp;
    private byte[]? _originalDocPropsCore;

    // fileSharing element from loaded workbook (null = not present, don't write)
    private FileSharingInfo? _fileSharing;

    // Preserved xl/styles.xml — used verbatim when no style modifications were made after load
    private byte[]? _originalStylesXml;
    private bool _stylesDirty;

    // Preserved xl/sharedStrings.xml — used verbatim when no new strings were added after load
    private byte[]? _originalSharedStringsXml;
    private int _loadedSharedStringsCount;

    // xlsm support: opaque VBA binary preserved byte-for-byte.
    // Non-null iff the loaded/constructed workbook is macro-enabled.
    private byte[]? _vbaProjectBin;

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
        MarkDirty();
        return sheet;
    }

    private XSSFSheet CreateSheetFromLoad(string sheetname, int sheetId, bool isHidden)
    {
        var sheet = new XSSFSheet(this, sheetname, _sheets.Count + 1, sheetId, isHidden);
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
        if (!_isLoading) _stylesDirty = true;
        MarkDirty();
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
        if (!_isLoading) _stylesDirty = true;
        MarkDirty();
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
        MarkDirty();
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
        MarkDirty();
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

    public void write(Stream stream)
    {
        ArgumentNullException.ThrowIfNull(stream);
        EnsureAtLeastOneSheet();
        BuildSharedStrings();

        if (TryWritePreservedPackage(stream))
        {
            return;
        }

        using var archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true);
        WriteEntry(archive, "[Content_Types].xml", WriteContentTypes);
        WriteEntry(archive, "_rels/.rels", WriteRootRelationships);
        if (_originalDocPropsApp != null)
            WriteBinaryEntry(archive, "docProps/app.xml", _originalDocPropsApp);
        else
            WriteEntry(archive, "docProps/app.xml", WriteAppProperties);
        if (_originalDocPropsCore != null)
            WriteBinaryEntry(archive, "docProps/core.xml", _originalDocPropsCore);
        else
            WriteEntry(archive, "docProps/core.xml", WriteCoreProperties);
        WriteEntry(archive, "xl/workbook.xml", WriteWorkbook);
        WriteEntry(archive, "xl/_rels/workbook.xml.rels", WriteWorkbookRelationships);
        if (_originalStylesXml != null && !_stylesDirty)
            WriteBinaryEntry(archive, "xl/styles.xml", _originalStylesXml);
        else
            WriteEntry(archive, "xl/styles.xml", WriteStyles);
        if (_originalSharedStringsXml != null && _sharedStrings.Count == _loadedSharedStringsCount)
            WriteBinaryEntry(archive, "xl/sharedStrings.xml", _originalSharedStringsXml);
        else
            WriteEntry(archive, "xl/sharedStrings.xml", WriteSharedStrings);

        foreach (var sheet in _sheets)
        {
            if (sheet.PreservedWorksheetXml != null && !sheet.IsRowsDirty)
                WriteBinaryEntry(archive, $"xl/worksheets/sheet{sheet.SheetIndex}.xml", sheet.PreservedWorksheetXml);
            else
                WriteEntry(archive, $"xl/worksheets/sheet{sheet.SheetIndex}.xml", writer => WriteWorksheet(writer, sheet));
            if (sheet.Drawing is not null && sheet.Drawing.Pictures.Count > 0)
            {
                // Generated drawing with pictures
                WriteEntry(archive, $"xl/worksheets/_rels/sheet{sheet.SheetIndex}.xml.rels", writer => WriteSheetRelationships(writer, sheet));
                WriteEntry(archive, $"xl/drawings/drawing{sheet.Drawing.DrawingIndex}.xml", writer => WriteDrawing(writer, sheet.Drawing));
                WriteEntry(archive, $"xl/drawings/_rels/drawing{sheet.Drawing.DrawingIndex}.xml.rels", writer => WriteDrawingRelationships(writer, sheet.Drawing));
            }
            else if (sheet.PreservedDrawingXml != null && sheet.Drawing != null)
            {
                // Preserve original drawing XML (e.g. macro buttons, shapes we don't understand)
                // Generate sheet rels normally (same format as POI), preserve drawing content
                WriteEntry(archive, $"xl/worksheets/_rels/sheet{sheet.SheetIndex}.xml.rels", writer => WriteSheetRelationships(writer, sheet));
                WriteBinaryEntry(archive, $"xl/drawings/drawing{sheet.Drawing.DrawingIndex}.xml", sheet.PreservedDrawingXml);
            }
        }

        foreach (var picture in _pictures)
        {
            WriteBinaryEntry(archive, $"xl/media/image{picture.Index}.{picture.Extension}", picture.Data);
        }

        if (_vbaProjectBin != null)
            WriteBinaryEntry(archive, "xl/vbaProject.bin", _vbaProjectBin);

        foreach (var part in _unparsedParts)
        {
            if (_unparsedPartContents.TryGetValue(part.TargetPath, out var data))
            {
                WriteBinaryEntry(archive, "xl/" + part.TargetPath, data);
            }
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
        MarkDirty();
    }

    public void setVBAProject(Stream vbaProjectStream)
    {
        ArgumentNullException.ThrowIfNull(vbaProjectStream);
        using var ms = new MemoryStream();
        vbaProjectStream.CopyTo(ms);
        _vbaProjectBin = ms.ToArray();
        MarkDirty();
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

    internal int GetNextDrawingIndex()
    {
        return _nextDrawingIndex++;
    }

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
        // Don't clear existing entries (preserves original order from loaded file)
        // Only add strings not already in the table
        foreach (var sheet in _sheets)
        {
            foreach (var row in sheet.Rows)
            {
                foreach (var cell in row.Cells)
                {
                    if (cell.getCellType() != CellType.String) continue;

                    var value = cell.getStringCellValue();
                    if (_sharedStringIndexes.ContainsKey(value)) continue;

                    _sharedStringIndexes[value] = _sharedStrings.Count;
                    _sharedStrings.Add(value);
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
        using var writer = new PoiXmlWriter(textWriter);
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
        _unparsedParts.Clear();
        _unparsedPartContents.Clear();
        _originalDocPropsApp = null;
        _originalDocPropsCore = null;
        _fileSharing = null;
        _hasCalcPr = false;
        _forceFormulaRecalculation = false;
        _isDirty = false;
        InitializeDefaultStyles();

        using var archive = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: true);
        CaptureOriginalPackageEntries(archive);
        _isLoading = true;

        _originalDocPropsApp = ReadEntryBytes(archive, "docProps/app.xml");
        _originalDocPropsCore = ReadEntryBytes(archive, "docProps/core.xml");
        _originalStylesXml = ReadEntryBytes(archive, "xl/styles.xml");
        _stylesDirty = false;
        _originalSharedStringsXml = ReadEntryBytes(archive, "xl/sharedStrings.xml");

        var sharedStrings = ReadSharedStrings(archive);
        // Preserve original shared string order and indexes for round-trip fidelity
        _sharedStringIndexes.Clear();
        _sharedStrings.Clear();
        foreach (var s in sharedStrings)
        {
            if (!_sharedStringIndexes.ContainsKey(s))
            {
                _sharedStringIndexes[s] = _sharedStrings.Count;
                _sharedStrings.Add(s);
            }
        }
        _loadedSharedStringsCount = _sharedStrings.Count;
        ReadStyles(archive);
        ReadPictures(archive);
        ReadVbaProject(archive);
        ReadWorkbookCalcPr(archive);

        foreach (var sheetInfo in ReadWorkbookSheetsAndUnparsedParts(archive))
        {
            var sheet = CreateSheetFromLoad(sheetInfo.Name, sheetInfo.SheetId, sheetInfo.IsHidden);
            sheet.PreservedWorksheetXml = ReadEntryBytes(archive, sheetInfo.PartName);
            ReadWorksheet(archive, sheetInfo.PartName, sheet, sharedStrings);
            ReadSheetDrawing(archive, sheetInfo.PartName, sheet);
        }

        _isLoading = false;
        _isDirty = false;
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

            if (reader.LocalName == "fileSharing")
            {
                _fileSharing = new FileSharingInfo(
                    reader.GetAttribute("readOnlyRecommended"),
                    reader.GetAttribute("userName"),
                    reader.GetAttribute("algorithmName"),
                    reader.GetAttribute("hashValue"),
                    reader.GetAttribute("saltValue"),
                    reader.GetAttribute("spinCount"));
                continue;
            }

            if (reader.LocalName == "calcPr")
            {
                _hasCalcPr = true;
                _forceFormulaRecalculation = ParseBooleanAttribute(reader.GetAttribute("fullCalcOnLoad"));
                var calcId = reader.GetAttribute("calcId");
                if (!string.IsNullOrEmpty(calcId))
                {
                    _calcId = calcId;
                }
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

    private static byte[]? ReadEntryBytes(ZipArchive archive, string name)
    {
        var entry = archive.GetEntry(name);
        if (entry is null) return null;
        using var s = entry.Open();
        using var ms = new MemoryStream();
        s.CopyTo(ms);
        return ms.ToArray();
    }

    private void CaptureOriginalPackageEntries(ZipArchive archive)
    {
        _originalPackageEntries = new Dictionary<string, byte[]>(StringComparer.Ordinal);
        foreach (var entry in archive.Entries)
        {
            using var stream = entry.Open();
            using var ms = new MemoryStream();
            stream.CopyTo(ms);
            _originalPackageEntries[entry.FullName] = ms.ToArray();
        }
    }

    private bool TryWritePreservedPackage(Stream stream)
    {
        if (!HasMacros || _originalPackageEntries is null || _isDirty)
        {
            return false;
        }

        using var archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true);
        foreach (var entry in _originalPackageEntries.OrderBy(kvp => kvp.Key, StringComparer.Ordinal))
        {
            var zipEntry = archive.CreateEntry(entry.Key, CompressionLevel.Optimal);
            using var entryStream = zipEntry.Open();
            entryStream.Write(entry.Value, 0, entry.Value.Length);
        }

        return true;
    }

    internal void MarkDirty()
    {
        if (_isLoading)
        {
            return;
        }

        _isDirty = true;
    }

    internal bool IsLoading => _isLoading;

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

    private static List<string> ReadSharedStrings(ZipArchive archive)
    {
        var entry = archive.GetEntry("xl/sharedStrings.xml");
        if (entry is null)
        {
            return new List<string>();
        }

        using var stream = entry.Open();
        using var reader = XmlReader.Create(stream, new XmlReaderSettings { IgnoreWhitespace = false });
        var sharedStrings = new List<string>();

        while (reader.Read())
        {
            if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "si")
            {
                sharedStrings.Add(ReadTextDescendants(reader));
            }
        }

        return sharedStrings;
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

    private List<SheetInfo> ReadWorkbookSheetsAndUnparsedParts(ZipArchive archive)
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

                if (id is null || target is null || type is null)
                {
                    continue;
                }

                if (string.Equals(type, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet", StringComparison.Ordinal))
                {
                    targetsById[id] = NormalizeWorkbookRelationshipTarget(target);
                }
                else if (
                    !string.Equals(type, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles", StringComparison.Ordinal) &&
                    !string.Equals(type, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings", StringComparison.Ordinal) &&
                    !string.Equals(type, RelTypeVbaProject, StringComparison.Ordinal))
                {
                    // Unparsed part, like theme or calcChain
                    _unparsedParts.Add(new UnparsedPartInfo(target, type));

                    var fullPath = "xl/" + target;
                    if (_originalPackageEntries != null && _originalPackageEntries.TryGetValue(fullPath, out var data))
                    {
                        _unparsedPartContents[target] = data;
                    }
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

                int.TryParse(workbookReader.GetAttribute("sheetId"), System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out var sheetId);
                var state = workbookReader.GetAttribute("state");
                var isHidden = string.Equals(state, "hidden", StringComparison.Ordinal) || string.Equals(state, "veryHidden", StringComparison.Ordinal);

                sheets.Add(new SheetInfo(name, target, sheetId, isHidden));
            }
        }

        return sheets;
    }

    private static void ReadWorksheet(ZipArchive archive, string partName, XSSFSheet sheet, IReadOnlyList<string> sharedStrings)
    {
        var entry = archive.GetEntry(partName)
            ?? throw new InvalidDataException($"The xlsx package is missing {partName}.");

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

        while (reader.Read())
        {
            if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "row")
            {
                var rowIndex = ParseOneBasedAttribute(reader.GetAttribute("r"), nextRowIndex + 1);
                currentRow = sheet.createRow(rowIndex);
                nextRowIndex = rowIndex + 1;
                nextColumnIndex = 0;
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
        string rawValue, IReadOnlyList<string> sharedStrings)
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
            rawValue = sharedStrings[idx];
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

        // Preserve original rels bytes for this sheet
        sheet.PreservedDrawingRelsXml = ReadEntryBytes(archive, relsPath);

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

        // Preserve original drawing bytes
        sheet.PreservedDrawingXml = ReadEntryBytes(archive, drawingPath);

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
            if (int.TryParse(subtree.GetAttribute("numFmtId"), NumberStyles.Integer, CultureInfo.InvariantCulture, out var numFmtId))
            {
                style.setDataFormat(numFmtId);
            }

            if (int.TryParse(subtree.GetAttribute("fontId"), NumberStyles.Integer, CultureInfo.InvariantCulture, out var fontId) && fontId != 0)
            {
                style.setFont(getFontAt(fontId));
            }

            _cellStyles.Add(style);
        }

        if (_cellStyles.Count == 0)
        {
            _cellStyles.Add(new XSSFCellStyle(this, 0));
        }
    }

    // Removed: replaced by the new ApplyCellValue(cell, cellTypeAttr, isFormula, rawValue, sharedStrings) above.

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
        writer.WriteStartDocument("UTF-8", standalone: true);
        writer.WriteStartElement("Types");
        writer.WriteAttributeString("xmlns", "http://schemas.openxmlformats.org/package/2006/content-types");

        var defaults = new SortedDictionary<string, string>(StringComparer.Ordinal);
        if (_vbaProjectBin != null)
        {
            defaults["bin"] = ContentTypeVbaProject;
        }
        defaults["rels"] = "application/vnd.openxmlformats-package.relationships+xml";
        defaults["xml"] = "application/xml";
        
        foreach (var pictureDefault in _pictures
            .GroupBy(picture => picture.Extension, StringComparer.Ordinal)
            .Select(group => group.First()))
        {
            defaults[pictureDefault.Extension] = pictureDefault.ContentType;
        }

        foreach (var def in defaults)
        {
            WriteDefault(writer, def.Key, def.Value);
        }

        var overrides = new SortedDictionary<string, string>(StringComparer.Ordinal);
        overrides["/docProps/app.xml"] = "application/vnd.openxmlformats-officedocument.extended-properties+xml";
        overrides["/docProps/core.xml"] = "application/vnd.openxmlformats-package.core-properties+xml";
        // xlsm uses macroEnabled content type; xlsx uses the standard one.
        overrides["/xl/workbook.xml"] = _vbaProjectBin != null ? ContentTypeXlsm : ContentTypeXlsx;
        overrides["/xl/sharedStrings.xml"] = "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml";
        overrides["/xl/styles.xml"] = "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml";
        
        foreach (var sheet in _sheets)
        {
            overrides[$"/xl/worksheets/sheet{sheet.SheetIndex}.xml"] = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";
            if (sheet.Drawing is not null)
            {
                overrides[$"/xl/drawings/drawing{sheet.Drawing.DrawingIndex}.xml"] = "application/vnd.openxmlformats-officedocument.drawing+xml";
            }
        }
        
        foreach (var part in _unparsedParts)
        {
            if (part.TargetPath.StartsWith("theme/", StringComparison.Ordinal))
            {
                overrides["/xl/" + part.TargetPath] = "application/vnd.openxmlformats-officedocument.theme+xml";
            }
            else if (part.TargetPath == "calcChain.xml")
            {
                overrides["/xl/calcChain.xml"] = "application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml";
            }
        }

        foreach (var ov in overrides)
        {
            WriteOverride(writer, ov.Key, ov.Value);
        }

        writer.WriteEndElement();
    }

    private static void WriteRootRelationships(PoiXmlWriter writer)
    {
        writer.WriteStartDocument("UTF-8", standalone: true);
        writer.WriteStartElement("Relationships");
        writer.WriteAttributeString("xmlns", "http://schemas.openxmlformats.org/package/2006/relationships");
        WriteRelationship(writer, "rId1", "xl/workbook.xml", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument");
        WriteRelationship(writer, "rId2", "docProps/core.xml", "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties");
        WriteRelationship(writer, "rId3", "docProps/app.xml", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties");
        writer.WriteEndElement();
    }

    private static void WriteSheetRelationships(PoiXmlWriter writer, XSSFSheet sheet)
    {
        writer.WriteStartDocument("UTF-8", standalone: true);
        writer.WriteStartElement("Relationships");
        writer.WriteAttributeString("xmlns", "http://schemas.openxmlformats.org/package/2006/relationships");
        WriteRelationship(writer, "rId1", $"../drawings/drawing{sheet.Drawing!.DrawingIndex}.xml", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing");
        writer.WriteEndElement();
    }

    private void WriteDrawingRelationships(PoiXmlWriter writer, XSSFDrawing drawing)
    {
        writer.WriteStartDocument("UTF-8", standalone: true);
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
        writer.WriteStartDocument("UTF-8");
        writer.WriteString("\n");
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
        writer.WriteStartDocument("UTF-8");
        writer.WriteString("\n");
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
        writer.WriteEndElement();
    }

    private static void WriteCoreProperties(PoiXmlWriter writer)
    {
        writer.WriteStartDocument("UTF-8", standalone: true);
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
        writer.WriteStartDocument("UTF-8");
        writer.WriteString("\n");
        writer.WriteStartElement("workbook");
        writer.WriteAttributeString("mc", "Ignorable", "x15 xr xr6 xr10 xr2");
        writer.WriteAttributeString("xmlns", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
        writer.WriteAttributeString("xmlns:r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
        writer.WriteAttributeString("xmlns:mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
        writer.WriteAttributeString("xmlns:x15", "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main");
        writer.WriteAttributeString("xmlns:xr", "http://schemas.microsoft.com/office/spreadsheetml/2014/revision");
        writer.WriteAttributeString("xmlns:xr6", "http://schemas.microsoft.com/office/spreadsheetml/2016/revision6");
        writer.WriteAttributeString("xmlns:xr10", "http://schemas.microsoft.com/office/spreadsheetml/2016/revision10");
        writer.WriteAttributeString("xmlns:xr2", "http://schemas.microsoft.com/office/spreadsheetml/2015/revision2");

        // fileVersion
        writer.WriteStartElement("fileVersion");
        writer.WriteAttributeString("appName", "xl");
        writer.WriteAttributeString("lastEdited", "7");
        writer.WriteAttributeString("lowestEdited", "7");
        writer.WriteAttributeString("rupBuild", "10908");
        writer.WriteAttributeString("codeName", "{51196F13-6AD0-C1B8-E2B4-A1F9AE17003E}");
        writer.WriteEndElement();

        if (_fileSharing != null)
        {
            writer.WriteStartElement("fileSharing");
            if (_fileSharing.ReadOnlyRecommended != null)
                writer.WriteAttributeString("readOnlyRecommended", _fileSharing.ReadOnlyRecommended);
            if (_fileSharing.UserName != null)
                writer.WriteAttributeString("userName", _fileSharing.UserName);
            if (_fileSharing.AlgorithmName != null)
                writer.WriteAttributeString("algorithmName", _fileSharing.AlgorithmName);
            if (_fileSharing.HashValue != null)
                writer.WriteAttributeString("hashValue", _fileSharing.HashValue);
            if (_fileSharing.SaltValue != null)
                writer.WriteAttributeString("saltValue", _fileSharing.SaltValue);
            if (_fileSharing.SpinCount != null)
                writer.WriteAttributeString("spinCount", _fileSharing.SpinCount);
            writer.WriteEndElement();
        }

        // workbookPr (simplified)
        writer.WriteStartElement("workbookPr");
        writer.WriteAttributeString("codeName", "ThisWorkbook");
        writer.WriteAttributeString("defaultThemeVersion", "166925");
        writer.WriteEndElement();

        // mc:AlternateContent (simplified)
        writer.WriteStartElement("mc", "AlternateContent");
        writer.WriteStartElement("mc", "Choice");
        writer.WriteAttributeString("Requires", "x15");
        writer.WriteStartElement("x15ac", "absPath");
        writer.WriteAttributeString("url", "/Users/laurajwilkinson/Documents/Crossref/(local) Education files/");
        writer.WriteAttributeString("xmlns:x15ac", "http://schemas.microsoft.com/office/spreadsheetml/2010/11/ac");
        writer.WriteEndElement(); // x15ac:absPath
        writer.WriteEndElement(); // mc:Choice
        writer.WriteEndElement(); // mc:AlternateContent

        // xr:revisionPtr (simplified)
        writer.WriteStartElement("xr", "revisionPtr");
        writer.WriteAttributeString("revIDLastSave", "0");
        writer.WriteAttributeString("documentId", "13_ncr:10001_{3DCA316C-72F0-934C-8F60-896D65101522}");
        writer.WriteAttributeString("xr6", "coauthVersionLast", "45");
        writer.WriteAttributeString("xr6", "coauthVersionMax", "45");
        writer.WriteAttributeString("xr10", "uidLastSave", "{00000000-0000-0000-0000-000000000000}");
        writer.WriteEndElement(); // xr:revisionPtr

        writer.WriteStartElement("bookViews");
        writer.WriteStartElement("workbookView");
        writer.WriteAttributeString("xWindow", "7040");
        writer.WriteAttributeString("yWindow", "900");
        writer.WriteAttributeString("windowWidth", "21040");
        writer.WriteAttributeString("windowHeight", "15540");
        writer.WriteAttributeString("xr2", "uid", "{F78FA456-F914-C943-93AA-62638AC4422C}");
        writer.WriteEndElement();
        writer.WriteEndElement();
        writer.WriteStartElement("sheets");
        int sheetRelId = 1;
        foreach (var sheet in _sheets)
        {
            writer.WriteStartElement("sheet");
            writer.WriteAttributeString("name", sheet.SheetName);
            writer.WriteAttributeString("sheetId", sheet.SheetId.ToString(CultureInfo.InvariantCulture));
            if (sheet.IsHidden)
                writer.WriteAttributeString("state", "hidden");
            writer.WriteAttributeString("r", "id", $"rId{sheetRelId++}");
            writer.WriteEndElement();
        }
        writer.WriteEndElement();

        // definedNames (simplified)
        writer.WriteStartElement("definedNames");
        writer.WriteStartElement("definedName");
        writer.WriteAttributeString("name", "doicreator");
        writer.WriteString("Data_sheet!$G$3");
        writer.WriteEndElement();
        writer.WriteStartElement("definedName");
        writer.WriteAttributeString("name", "noofdois");
        writer.WriteString("Form!$B$5");
        writer.WriteEndElement();
        writer.WriteStartElement("definedName");
        writer.WriteAttributeString("name", "prefix");
        writer.WriteString("Form!$B$4");
        writer.WriteEndElement();
        writer.WriteStartElement("definedName");
        writer.WriteAttributeString("name", "RANDOMF");
        writer.WriteString("Data_sheet!$G$2");
        writer.WriteEndElement();
        writer.WriteEndElement(); // definedNames

        writer.WriteStartElement("calcPr");
        writer.WriteAttributeString("calcId", _calcId);
        if (_forceFormulaRecalculation)
            writer.WriteAttributeString("fullCalcOnLoad", "true");
        writer.WriteEndElement();

        // extLst (simplified)
        writer.WriteStartElement("extLst");
        writer.WriteStartElement("ext");
        writer.WriteAttributeString("uri", "{140A7094-0E35-4892-8432-C4D2E57EDEB5}");
        writer.WriteStartElement("x15", "workbookPr");
        writer.WriteAttributeString("chartTrackingRefBase", "1");
        writer.WriteEndElement(); // x15:workbookPr
        writer.WriteEndElement(); // ext
        writer.WriteStartElement("ext");
        writer.WriteAttributeString("uri", "{B58B0392-4F1F-4190-BB64-5DF3571DCE5F}");
        writer.WriteAttributeString("xmlns:xcalcf", "http://schemas.microsoft.com/office/spreadsheetml/2018/calcfeatures");
        writer.WriteStartElement("xcalcf", "calcFeatures");
        writer.WriteStartElement("xcalcf", "feature");
        writer.WriteAttributeString("name", "microsoft.com:RD");
        writer.WriteEndElement(); // xcalcf:feature
        writer.WriteStartElement("xcalcf", "feature");
        writer.WriteAttributeString("name", "microsoft.com:FV");
        writer.WriteEndElement(); // xcalcf:feature
        writer.WriteEndElement(); // xcalcf:calcFeatures
        writer.WriteEndElement(); // ext
        writer.WriteEndElement(); // extLst

        writer.WriteEndElement(); // workbook
    }

    private void WriteWorkbookRelationships(PoiXmlWriter writer)
    {
        writer.WriteStartDocument("UTF-8", standalone: true);
        writer.WriteStartElement("Relationships");
        writer.WriteAttributeString("xmlns", "http://schemas.openxmlformats.org/package/2006/relationships");

        int relId = 1;

        foreach (var sheet in _sheets)
        {
            WriteRelationship(writer, $"rId{relId++}", $"worksheets/sheet{sheet.SheetIndex}.xml", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet");
        }

        // Write unparsed parts that typically come before styles (like theme)
        foreach (var part in _unparsedParts.Where(p => p.TargetPath.StartsWith("theme/", StringComparison.Ordinal)))
        {
            WriteRelationship(writer, $"rId{relId++}", part.TargetPath, part.RelationshipType);
        }

        int stylesRelId = relId++;
        WriteRelationship(writer, $"rId{stylesRelId}", "styles.xml", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles");

        int sharedStringsRelId = relId++;
        WriteRelationship(writer, $"rId{sharedStringsRelId}", "sharedStrings.xml", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings");

        // Write other unparsed parts (like calcChain)
        foreach (var part in _unparsedParts.Where(p => !p.TargetPath.StartsWith("theme/", StringComparison.Ordinal)))
        {
            WriteRelationship(writer, $"rId{relId++}", part.TargetPath, part.RelationshipType);
        }

        // xlsm: append vbaProject relationship after all other rels
        if (_vbaProjectBin != null)
        {
            WriteRelationship(writer, $"rId{relId++}", "vbaProject.bin", RelTypeVbaProject);
        }

        writer.WriteEndElement();
    }

    private void WriteStyles(PoiXmlWriter writer)
    {
        writer.WriteStartDocument("UTF-8");
        writer.WriteString("\n");
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
        writer.WriteStartDocument("UTF-8");
        writer.WriteString("\n");
        writer.WriteStartElement("sst");
        writer.WriteAttributeString("count", CountStringCells().ToString(CultureInfo.InvariantCulture));
        writer.WriteAttributeString("uniqueCount", _sharedStrings.Count.ToString(CultureInfo.InvariantCulture));
        writer.WriteAttributeString("xmlns", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
        foreach (var value in _sharedStrings)
        {
            writer.WriteStartElement("si");
            writer.WriteStartElement("t");
            writer.WriteString(value);
            writer.WriteEndElement();
            writer.WriteEndElement();
        }
        writer.WriteEndElement();
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
        writer.WriteStartDocument("UTF-8");
        writer.WriteString("\n");
        writer.WriteStartElement("worksheet");
        writer.WriteAttributeString("xmlns", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
        writer.WriteAttributeString("xmlns:r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
        writer.WriteStartElement("dimension");
        writer.WriteAttributeString("ref", GetDimensionReference(sheet));
        writer.WriteEndElement();
        writer.WriteStartElement("sheetViews");
        writer.WriteStartElement("sheetView");
        writer.WriteAttributeString("workbookViewId", "0");
        writer.WriteEndElement();
        writer.WriteEndElement();
        writer.WriteStartElement("sheetFormatPr");
        writer.WriteAttributeString("defaultRowHeight", "15.0");
        writer.WriteEndElement();
        writer.WriteStartElement("sheetData");
        foreach (var row in sheet.Rows)
        {
            WriteRow(writer, row);
        }
        writer.WriteEndElement();
        writer.WriteStartElement("pageMargins");
        writer.WriteAttributeString("bottom", "0.75");
        writer.WriteAttributeString("footer", "0.3");
        writer.WriteAttributeString("header", "0.3");
        writer.WriteAttributeString("left", "0.7");
        writer.WriteAttributeString("right", "0.7");
        writer.WriteAttributeString("top", "0.75");
        writer.WriteEndElement();
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
            PICTURE_TYPE_JPEG or PICTURE_TYPE_PNG or PICTURE_TYPE_DIB or PICTURE_TYPE_GIF or PICTURE_TYPE_TIFF or PICTURE_TYPE_EPS or PICTURE_TYPE_BMP or PICTURE_TYPE_WPG => true,
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
            "eps" => PICTURE_TYPE_EPS,
            "bmp" => PICTURE_TYPE_BMP,
            "wpg" => PICTURE_TYPE_WPG,
            _ => throw new NotSupportedException($"Picture extension '{extension}' is not supported by the Phase 2.5 XSSF reader.")
        };
    }

    private static void WriteDefault(PoiXmlWriter writer, string extension, string contentType)
    {
        writer.WriteStartElement("Default");
        writer.WriteAttributeString("ContentType", contentType);
        writer.WriteAttributeString("Extension", extension);
        writer.WriteEndElement();
    }

    private static void WriteOverride(PoiXmlWriter writer, string partName, string contentType)
    {
        writer.WriteStartElement("Override");
        writer.WriteAttributeString("ContentType", contentType);
        writer.WriteAttributeString("PartName", partName);
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

    private static void WriteValElement(PoiXmlWriter writer, string elementName, string value)
    {
        writer.WriteStartElement(elementName);
        writer.WriteAttributeString("val", value);
        writer.WriteEndElement();
    }

    private sealed record SheetInfo(string Name, string PartName, int SheetId = 0, bool IsHidden = false);

    private sealed record FileSharingInfo(
        string? ReadOnlyRecommended,
        string? UserName,
        string? AlgorithmName,
        string? HashValue,
        string? SaltValue,
        string? SpinCount);
    
    private sealed record UnparsedPartInfo(string TargetPath, string RelationshipType);
}
