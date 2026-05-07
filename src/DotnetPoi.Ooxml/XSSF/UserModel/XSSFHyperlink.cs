using DotnetPoi.SS.UserModel;

namespace DotnetPoi.XSSF.UserModel;

/// <summary>
/// XSSF-specific implementation of IHyperlink.
/// Ported from org.apache.poi.xssf.usermodel.XSSFHyperlink.
/// </summary>
public sealed class XSSFHyperlink : IHyperlink
{
    public string Address { get; set; } = string.Empty;
    public string CellRef { get; set; } = string.Empty;
    public HyperlinkType Type { get; }

    /// <summary>The relationship ID used in the sheet rels file (e.g., "rId1").</summary>
    internal string? RelationshipId { get; set; }

    /// <summary>True if the target is external (URL/file), false for internal workbook links.</summary>
    internal bool IsExternal { get; set; }

    public XSSFHyperlink(HyperlinkType type)
    {
        Type = type;
        IsExternal = type == HyperlinkType.Url || type == HyperlinkType.File || type == HyperlinkType.Email;
    }

    public string getAddress() => Address;

    public void setAddress(string address)
    {
        Address = address ?? string.Empty;
    }

    public string getCellRef() => CellRef;

    public void setCellRef(string cellRef)
    {
        CellRef = cellRef ?? string.Empty;
    }

    public HyperlinkType getType() => Type;

    /// <summary>Creates a cell reference string from row/column indices (0-based).</summary>
    internal static string FormatCellRef(int row, int col)
    {
        return FormatColumn(col) + (row + 1);
    }

    private static string FormatColumn(int col)
    {
        var result = new System.Text.StringBuilder();
        col++; // 1-based
        while (col > 0)
        {
            col--;
            result.Insert(0, (char)('A' + (col % 26)));
            col /= 26;
        }
        return result.ToString();
    }
}
