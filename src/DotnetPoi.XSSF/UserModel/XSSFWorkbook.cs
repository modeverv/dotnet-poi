using System.Globalization;
using System.IO.Compression;
using System.Text;
using DotnetPoi.SS.Xml;

namespace DotnetPoi.XSSF.UserModel;

public sealed class XSSFWorkbook : IDisposable
{
    private readonly List<XSSFSheet> _sheets = new();
    private readonly Dictionary<string, int> _sharedStringIndexes = new(StringComparer.Ordinal);
    private readonly List<string> _sharedStrings = new();
    private XSSFCreationHelper? _creationHelper;

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

    public int getNumberOfSheets()
    {
        return _sheets.Count;
    }

    public XSSFCreationHelper getCreationHelper()
    {
        return _creationHelper ??= new XSSFCreationHelper(this);
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

    private void WriteContentTypes(PoiXmlWriter writer)
    {
        writer.WriteStartDocument("UTF-8", standalone: true);
        writer.WriteStartElement("Types");
        writer.WriteAttributeString("xmlns", "http://schemas.openxmlformats.org/package/2006/content-types");
        WriteDefault(writer, "rels", "application/vnd.openxmlformats-package.relationships+xml");
        WriteDefault(writer, "xml", "application/xml");
        WriteOverride(writer, "/docProps/app.xml", "application/vnd.openxmlformats-officedocument.extended-properties+xml");
        WriteOverride(writer, "/docProps/core.xml", "application/vnd.openxmlformats-package.core-properties+xml");
        WriteOverride(writer, "/xl/workbook.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml");
        WriteOverride(writer, "/xl/sharedStrings.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml");
        WriteOverride(writer, "/xl/styles.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml");
        foreach (var sheet in _sheets)
        {
            WriteOverride(writer, $"/xl/worksheets/sheet{sheet.SheetIndex}.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml");
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

    private static void WriteStyles(PoiXmlWriter writer)
    {
        writer.WriteStartDocument("UTF-8", standalone: false);
        writer.WriteString("\n");
        writer.WriteStartElement("styleSheet");
        writer.WriteAttributeString("xmlns", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
        writer.WriteStartElement("numFmts");
        writer.WriteAttributeString("count", "0");
        writer.WriteEndElement();
        writer.WriteStartElement("fonts");
        writer.WriteAttributeString("count", "1");
        writer.WriteStartElement("font");
        WriteValElement(writer, "sz", "11.0");
        writer.WriteStartElement("color");
        writer.WriteAttributeString("indexed", "8");
        writer.WriteEndElement();
        WriteValElement(writer, "name", "Calibri");
        WriteValElement(writer, "family", "2");
        WriteValElement(writer, "scheme", "minor");
        writer.WriteEndElement();
        writer.WriteEndElement();
        writer.WriteStartElement("fills");
        writer.WriteAttributeString("count", "2");
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
        writer.WriteEndElement();
        writer.WriteStartElement("borders");
        writer.WriteAttributeString("count", "1");
        writer.WriteStartElement("border");
        writer.WriteStartElement("left");
        writer.WriteEndElement();
        writer.WriteStartElement("right");
        writer.WriteEndElement();
        writer.WriteStartElement("top");
        writer.WriteEndElement();
        writer.WriteStartElement("bottom");
        writer.WriteEndElement();
        writer.WriteStartElement("diagonal");
        writer.WriteEndElement();
        writer.WriteEndElement();
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
        writer.WriteAttributeString("count", "1");
        writer.WriteStartElement("xf");
        writer.WriteAttributeString("numFmtId", "0");
        writer.WriteAttributeString("fontId", "0");
        writer.WriteAttributeString("fillId", "0");
        writer.WriteAttributeString("borderId", "0");
        writer.WriteAttributeString("xfId", "0");
        writer.WriteEndElement();
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
}
