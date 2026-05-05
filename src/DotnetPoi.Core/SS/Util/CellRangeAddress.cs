using System.Globalization;

namespace DotnetPoi.SS.Util;

/// <summary>
/// Represents a range of cells in a sheet: (firstRow, lastRow, firstCol, lastCol).
/// Ported from org.apache.poi.ss.util.CellRangeAddress.
/// </summary>
public sealed class CellRangeAddress
{
    public int FirstRow { get; }
    public int LastRow { get; }
    public int FirstCol { get; }
    public int LastCol { get; }

    public CellRangeAddress(int firstRow, int lastRow, int firstCol, int lastCol)
    {
        FirstRow = firstRow;
        LastRow = lastRow;
        FirstCol = firstCol;
        LastCol = lastCol;
    }

    /// <summary>Parses an OOXML range reference like "A1:B2" or "$A$1:$C$3".</summary>
    public static CellRangeAddress Parse(string refText)
    {
        ArgumentNullException.ThrowIfNull(refText);

        var parts = refText.Split(':');
        if (parts.Length != 2)
            throw new ArgumentException($"Invalid cell range reference: '{refText}'.", nameof(refText));

        var (firstCol, firstRow) = ParseCellReference(parts[0].Trim());
        var (lastCol, lastRow) = ParseCellReference(parts[1].Trim());

        return new CellRangeAddress(firstRow, lastRow, firstCol, lastCol);
    }

    /// <summary>Formats this range as an OOXML range reference like "A1:B2".</summary>
    public string FormatAsString()
    {
        return FormatCellReference(FirstCol, FirstRow)
             + ":"
             + FormatCellReference(LastCol, LastRow);
    }

    private static (int col, int row) ParseCellReference(string reference)
    {
        // Strip any '$' prefix
        var clean = reference.TrimStart('$');
        int i = 0;
        while (i < clean.Length && char.IsLetter(clean[i])) i++;
        var colPart = clean.Substring(0, i);
        var rowPart = clean.Substring(i);

        return (ParseColumnLetters(colPart), int.Parse(rowPart, CultureInfo.InvariantCulture) - 1);
    }

    private static int ParseColumnLetters(string letters)
    {
        int col = 0;
        foreach (var ch in letters)
        {
            col = col * 26 + (char.ToUpperInvariant(ch) - 'A' + 1);
        }
        return col - 1;
    }

    private static string FormatCellReference(int col, int row)
    {
        return FormatColumnLetters(col) + (row + 1).ToString(CultureInfo.InvariantCulture);
    }

    private static string FormatColumnLetters(int col)
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
