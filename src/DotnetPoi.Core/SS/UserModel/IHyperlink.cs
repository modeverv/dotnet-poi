namespace DotnetPoi.SS.UserModel;

/// <summary>
/// Represents a hyperlink on a cell.
/// Ported from org.apache.poi.ss.usermodel.Hyperlink.
/// </summary>
public interface IHyperlink
{
    /// <summary>Gets the hyperlink address (URL, file path, email, etc.).</summary>
    string getAddress();

    /// <summary>Sets the hyperlink address.</summary>
    void setAddress(string address);

    /// <summary>Gets the cell reference this hyperlink is anchored to (e.g., "A1").</summary>
    string getCellRef();

    /// <summary>Sets the cell reference this hyperlink is anchored to.</summary>
    void setCellRef(string cellRef);

    /// <summary>Gets the type of hyperlink (e.g., Url, File, Email, Document).</summary>
    HyperlinkType getType();
}

/// <summary>
/// Enumeration of hyperlink types.
/// Ported from org.apache.poi.common.usermodel.HyperlinkType.
/// </summary>
public enum HyperlinkType
{
    None = 0,
    Url = 1,
    File = 2,
    Email = 3,
    Document = 4  // Internal link within the workbook
}
