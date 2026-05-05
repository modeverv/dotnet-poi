using System.Globalization;
using DotnetPoi.SS.Xml;

namespace DotnetPoi.XSSF.UserModel;

/// <summary>
/// Represents a pivot table definition part (xl/pivotTables/pivotTable{id}.xml).
/// Ported from Apache POI XSSFPivotTable.
/// Supports minimal creation: source area, destination cell, row labels, column labels, data fields.
/// </summary>
public sealed class XSSFPivotTable
{
    // OOXML constants used in pivotTableDefinition output
    private const string NsSpreadsheetMl = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private const string NsR = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    /// <summary>
    /// The 1-based pivot table index (determines part file name).
    /// </summary>
    public int PivotTableIndex { get; set; }

    /// <summary>
    /// The cache ID referenced from the pivot table definition.
    /// </summary>
    public int CacheId { get; set; }

    /// <summary>
    /// Display name of the pivot table.
    /// </summary>
    public string Name { get; set; } = string.Empty;

    /// <summary>
    /// Caption for the data area ("Values" is default).
    /// </summary>
    public string DataCaption { get; set; } = "Values";

    /// <summary>
    /// Destination cell reference for the pivot table location (e.g. "E5").
    /// </summary>
    public string? DestinationCell { get; set; }

    /// <summary>
    /// Source area reference (e.g. "Sheet1!A1:C100").
    /// </summary>
    public string? SourceAreaRef { get; set; }

    /// <summary>
    /// Source sheet name.
    /// </summary>
    public string? SourceSheetName { get; set; }

    /// <summary>
    /// Row label column indices (0-based, relative to source first column).
    /// </summary>
    public List<int> RowLabels { get; } = new();

    /// <summary>
    /// Column label column indices (0-based, relative to source first column).
    /// </summary>
    public List<int> ColumnLabels { get; } = new();

    /// <summary>
    /// Data field column indices (0-based, relative to source first column).
    /// </summary>
    public List<int> DataColumns { get; } = new();

    /// <summary>
    /// Associated cache definition for this pivot table.
    /// </summary>
    internal XSSFPivotCacheDefinition CacheDefinition { get; set; } = new();

    /// <summary>
    /// Associated cache records for this pivot table.
    /// </summary>
    internal XSSFPivotCacheRecords CacheRecords { get; set; } = new();

    /// <summary>
    /// Associated cache reference for this pivot table.
    /// </summary>
    internal XSSFPivotCache Cache { get; set; } = new();

    /// <summary>
    /// Relationship ID for pivotTable → pivotCacheDefinition link (in workbook.xml.rels).
    /// </summary>
    internal string? CacheRelId { get; set; }

    public XSSFPivotTable()
    {
    }

    /// <summary>
    /// Writes the pivotTableDefinition XML to the given writer.
    /// </summary>
    internal void WritePivotTableDefinition(PoiXmlWriter writer)
    {
        writer.WriteStartElement("pivotTableDefinition", NsSpreadsheetMl);
        writer.WriteAttributeString("xmlns:r", NsR);
        writer.WriteAttributeString("name", Name);
        writer.WriteAttributeString("cacheId", CacheId.ToString(CultureInfo.InvariantCulture));
        writer.WriteAttributeString("dataCaption", DataCaption);
        writer.WriteAttributeString("multipleFieldFilters", "0");
        writer.WriteAttributeString("indent", "0");
        writer.WriteAttributeString("createdVersion", "3");
        writer.WriteAttributeString("minRefreshableVersion", "3");
        writer.WriteAttributeString("updatedVersion", "3");
        writer.WriteAttributeString("itemPrintTitles", "1");
        writer.WriteAttributeString("useAutoFormatting", "1");
        writer.WriteAttributeString("applyNumberFormats", "0");
        writer.WriteAttributeString("applyWidthHeightFormats", "1");
        writer.WriteAttributeString("applyAlignmentFormats", "0");
        writer.WriteAttributeString("applyPatternFormats", "0");
        writer.WriteAttributeString("applyFontFormats", "0");
        writer.WriteAttributeString("applyBorderFormats", "0");

        // Location
        if (DestinationCell is not null)
        {
            // Compute a 2x2 destination area
            var (col, row) = ParseCellRefSimple(DestinationCell);
            var destEndCol = col + 1;
            var destEndRow = row + 1;
            var destRef = FormatColumnLetter(col) + (row + 1) + ":" + FormatColumnLetter(destEndCol) + (destEndRow + 1);

            writer.WriteStartElement("location");
            writer.WriteAttributeString("ref", destRef);
            writer.WriteAttributeString("firstHeaderRow", "1");
            writer.WriteAttributeString("firstDataRow", "1");
            writer.WriteAttributeString("firstDataCol", "1");
            writer.WriteEndElement();
        }

        // PivotFields — one per source column
        int sourceColumnCount = 0;
        if (SourceAreaRef is not null)
        {
            var parts = SourceAreaRef.Split(':');
            if (parts.Length == 2)
            {
                var (_, startRow) = ParseCellRefSimple(parts[0]);
                var (endCol, _) = ParseCellRefSimple(parts[1]);
                sourceColumnCount = endCol + 1;
            }
        }

        if (sourceColumnCount > 0)
        {
            writer.WriteStartElement("pivotFields");
            writer.WriteAttributeString("count", sourceColumnCount.ToString(CultureInfo.InvariantCulture));
            for (int i = 0; i < sourceColumnCount; i++)
            {
                bool isData = DataColumns.Contains(i);
                writer.WriteStartElement("pivotField");
                writer.WriteAttributeString("showAll", "0");
                if (isData)
                    writer.WriteAttributeString("dataField", "1");
                if (RowLabels.Contains(i))
                    writer.WriteAttributeString("axis", "axisRow");
                else if (ColumnLabels.Contains(i))
                    writer.WriteAttributeString("axis", "axisCol");
                writer.WriteEndElement();
            }
            writer.WriteEndElement();
        }

        // RowFields
        if (RowLabels.Count > 0)
        {
            writer.WriteStartElement("rowFields");
            writer.WriteAttributeString("count", RowLabels.Count.ToString(CultureInfo.InvariantCulture));
            foreach (var x in RowLabels)
            {
                writer.WriteStartElement("field");
                writer.WriteAttributeString("x", x.ToString(CultureInfo.InvariantCulture));
                writer.WriteEndElement();
            }
            writer.WriteEndElement();
        }

        // ColFields
        if (ColumnLabels.Count > 0)
        {
            writer.WriteStartElement("colFields");
            writer.WriteAttributeString("count", ColumnLabels.Count.ToString(CultureInfo.InvariantCulture));
            foreach (var x in ColumnLabels)
            {
                writer.WriteStartElement("field");
                writer.WriteAttributeString("x", x.ToString(CultureInfo.InvariantCulture));
                writer.WriteEndElement();
            }
            writer.WriteEndElement();
        }

        // DataFields
        if (DataColumns.Count > 0)
        {
            writer.WriteStartElement("dataFields");
            writer.WriteAttributeString("count", DataColumns.Count.ToString(CultureInfo.InvariantCulture));
            foreach (var x in DataColumns)
            {
                writer.WriteStartElement("dataField");
                writer.WriteAttributeString("name", "Sum of Field" + (x + 1).ToString(CultureInfo.InvariantCulture));
                writer.WriteAttributeString("fld", x.ToString(CultureInfo.InvariantCulture));
                writer.WriteAttributeString("subtotal", "sum");
                writer.WriteEndElement();
            }
            writer.WriteEndElement();
        }

        // Style info
        writer.WriteStartElement("pivotTableStyleInfo");
        writer.WriteAttributeString("name", "PivotStyleLight16");
        writer.WriteAttributeString("showLastColumn", "1");
        writer.WriteAttributeString("showColStripes", "0");
        writer.WriteAttributeString("showRowStripes", "0");
        writer.WriteAttributeString("showColHeaders", "1");
        writer.WriteAttributeString("showRowHeaders", "1");
        writer.WriteEndElement();

        writer.WriteEndElement(); // pivotTableDefinition
    }

    /// <summary>
    /// Parses a cell reference like "E5" into (col, row) 0-based.
    /// </summary>
    private static (int col, int row) ParseCellRefSimple(string reference)
    {
        var clean = reference.TrimStart('$');
        int i = 0;
        while (i < clean.Length && char.IsLetter(clean[i])) i++;
        var colPart = clean.Substring(0, i);
        var rowPart = clean.Substring(i);
        int col = 0;
        foreach (var ch in colPart)
            col = col * 26 + (char.ToUpperInvariant(ch) - 'A' + 1);
        return (col - 1, int.Parse(rowPart, CultureInfo.InvariantCulture) - 1);
    }

    /// <summary>
    /// Formats a 0-based column index as a column letter (e.g. 0 → "A", 26 → "AA").
    /// </summary>
    private static string FormatColumnLetter(int col)
    {
        var sb = new System.Text.StringBuilder();
        col++; // 1-based
        while (col > 0)
        {
            col--;
            sb.Insert(0, (char)('A' + col % 26));
            col /= 26;
        }
        return sb.ToString();
    }
}
