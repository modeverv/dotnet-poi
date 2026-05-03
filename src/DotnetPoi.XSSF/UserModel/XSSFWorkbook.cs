using System.Globalization;
using System.IO.Compression;
using System.Text;
using System.Xml;
using DotnetPoi.SS.Xml;

namespace DotnetPoi.XSSF.UserModel;

public sealed class XSSFWorkbook : IDisposable
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

    public void write(Stream stream)
    {
        ArgumentNullException.ThrowIfNull(stream);
        EnsureAtLeastOneSheet();
        BuildSharedStrings();

        using var archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true);
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
            if (sheet.Drawing is not null)
            {
                WriteEntry(archive, $"xl/worksheets/_rels/sheet{sheet.SheetIndex}.xml.rels", writer => WriteSheetRelationships(writer, sheet));
                WriteEntry(archive, $"xl/drawings/drawing{sheet.Drawing.DrawingIndex}.xml", writer => WriteDrawing(writer, sheet.Drawing));
                WriteEntry(archive, $"xl/drawings/_rels/drawing{sheet.Drawing.DrawingIndex}.xml.rels", writer => WriteDrawingRelationships(writer, sheet.Drawing));
            }
        }

        foreach (var picture in _pictures)
        {
            WriteBinaryEntry(archive, $"xl/media/image{picture.Index}.{picture.Extension}", picture.Data);
        }
    }

    public void close()
    {
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
        _sharedStringIndexes.Clear();
        _sharedStrings.Clear();

        foreach (var sheet in _sheets)
        {
            foreach (var row in sheet.Rows)
            {
                foreach (var cell in row.Cells)
                {
                    if (cell.getCellType() != XSSFCellType.String)
                    {
                        continue;
                    }

                    var value = cell.getStringCellValue();
                    if (_sharedStringIndexes.ContainsKey(value))
                    {
                        continue;
                    }

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
        InitializeDefaultStyles();

        using var archive = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: true);
        var sharedStrings = ReadSharedStrings(archive);
        ReadStyles(archive);
        ReadPictures(archive);

        foreach (var sheetInfo in ReadWorkbookSheets(archive))
        {
            var sheet = createSheet(sheetInfo.Name);
            ReadWorksheet(archive, sheetInfo.PartName, sheet, sharedStrings);
        }
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

    private static void ReadWorksheet(ZipArchive archive, string partName, XSSFSheet sheet, IReadOnlyList<string> sharedStrings)
    {
        var entry = archive.GetEntry(partName)
            ?? throw new InvalidDataException($"The xlsx package is missing {partName}.");

        using var stream = entry.Open();
        using var reader = XmlReader.Create(stream, new XmlReaderSettings { IgnoreWhitespace = false });
        XSSFRow? currentRow = null;
        XSSFCell? currentCell = null;
        string? currentCellType = null;
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
                currentCellType = reader.GetAttribute("t");
                var styleIndexText = reader.GetAttribute("s");
                if (styleIndexText is not null)
                {
                    currentCell.SetCellStyleIndex(int.Parse(styleIndexText, CultureInfo.InvariantCulture));
                }
                inlineText.Clear();
                nextColumnIndex = columnIndex + 1;

                if (reader.IsEmptyElement)
                {
                    currentCell = null;
                    currentCellType = null;
                }
                continue;
            }

            if (currentCell is not null && reader.NodeType == XmlNodeType.Element && reader.LocalName == "v")
            {
                ApplyCellValue(currentCell, currentCellType, reader.ReadElementContentAsString(), sharedStrings);
                continue;
            }

            if (currentCell is not null && currentCellType == "inlineStr" && reader.NodeType == XmlNodeType.Element && reader.LocalName == "t")
            {
                inlineText.Append(reader.ReadElementContentAsString());
                continue;
            }

            if (currentCell is not null && reader.NodeType == XmlNodeType.EndElement && reader.LocalName == "c")
            {
                if (currentCellType == "inlineStr")
                {
                    currentCell.setCellValue(inlineText.ToString());
                }

                currentCell = null;
                currentCellType = null;
            }
        }
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

    private static void ApplyCellValue(XSSFCell cell, string? cellType, string valueText, IReadOnlyList<string> sharedStrings)
    {
        if (cellType == "s")
        {
            var sharedStringIndex = int.Parse(valueText, CultureInfo.InvariantCulture);
            if ((uint)sharedStringIndex >= (uint)sharedStrings.Count)
            {
                throw new InvalidDataException($"Shared string index {sharedStringIndex} is outside the shared string table.");
            }

            cell.setCellValue(sharedStrings[sharedStringIndex]);
        }
        else
        {
            cell.setCellValue(double.Parse(valueText, CultureInfo.InvariantCulture));
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
        writer.WriteStartDocument("UTF-8", standalone: true);
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
        WriteOverride(writer, "/xl/workbook.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml");
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
        writer.WriteEndElement();
    }

    private static void WriteRootRelationships(PoiXmlWriter writer)
    {
        writer.WriteStartDocument("UTF-8", standalone: true);
        writer.WriteStartElement("Relationships");
        writer.WriteAttributeString("xmlns", "http://schemas.openxmlformats.org/package/2006/relationships");
        WriteRelationship(writer, "rId1", "xl/workbook.xml", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument");
        WriteRelationship(writer, "rId2", "docProps/app.xml", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties");
        WriteRelationship(writer, "rId3", "docProps/core.xml", "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties");
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
        writer.WriteStartDocument("UTF-8", standalone: false);
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
        writer.WriteStartDocument("UTF-8", standalone: false);
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
        writer.WriteStartDocument("UTF-8", standalone: false);
        writer.WriteString("\n");
        writer.WriteStartElement("workbook");
        writer.WriteAttributeString("xmlns", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
        writer.WriteAttributeString("xmlns:r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
        writer.WriteStartElement("workbookPr");
        writer.WriteAttributeString("date1904", "false");
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
        writer.WriteEndElement();
        writer.WriteEndElement();
    }

    private void WriteWorkbookRelationships(PoiXmlWriter writer)
    {
        writer.WriteStartDocument("UTF-8", standalone: true);
        writer.WriteStartElement("Relationships");
        writer.WriteAttributeString("xmlns", "http://schemas.openxmlformats.org/package/2006/relationships");
        WriteRelationship(writer, "rId1", "sharedStrings.xml", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings");
        WriteRelationship(writer, "rId2", "styles.xml", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles");
        foreach (var sheet in _sheets)
        {
            WriteRelationship(writer, "rId" + (sheet.SheetIndex + 2).ToString(CultureInfo.InvariantCulture), $"worksheets/sheet{sheet.SheetIndex}.xml", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet");
        }
        writer.WriteEndElement();
    }

    private void WriteStyles(PoiXmlWriter writer)
    {
        writer.WriteStartDocument("UTF-8", standalone: false);
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
        writer.WriteStartDocument("UTF-8", standalone: false);
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
                count += row.Cells.Count(cell => cell.getCellType() == XSSFCellType.String);
            }
        }

        return count;
    }

    private void WriteWorksheet(PoiXmlWriter writer, XSSFSheet sheet)
    {
        writer.WriteStartDocument("UTF-8", standalone: false);
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
            if (cell.getCellType() == XSSFCellType.Blank)
            {
                continue;
            }

            writer.WriteStartElement("c");
            writer.WriteAttributeString("r", FormatCellReference(cell.getRowIndex(), cell.getColumnIndex()));
            var styleIndex = cell.getCellStyle().getIndex();
            if (styleIndex != 0)
            {
                writer.WriteAttributeString("s", styleIndex.ToString(CultureInfo.InvariantCulture));
            }
            if (cell.getCellType() == XSSFCellType.String)
            {
                writer.WriteAttributeString("t", "s");
            }

            writer.WriteStartElement("v");
            if (cell.getCellType() == XSSFCellType.String)
            {
                writer.WriteString(GetSharedStringIndex(cell.getStringCellValue()).ToString(CultureInfo.InvariantCulture));
            }
            else
            {
                writer.WriteString(cell.GetNumericText());
            }
            writer.WriteEndElement();
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
                if (cell.getCellType() == XSSFCellType.Blank)
                {
                    continue;
                }

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

    private static void WriteValElement(PoiXmlWriter writer, string elementName, string value)
    {
        writer.WriteStartElement(elementName);
        writer.WriteAttributeString("val", value);
        writer.WriteEndElement();
    }

    private sealed record SheetInfo(string Name, string PartName);
}
