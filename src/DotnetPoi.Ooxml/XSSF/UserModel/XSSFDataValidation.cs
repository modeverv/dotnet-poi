using DotnetPoi.SS.UserModel;
using DotnetPoi.SS.Util;

namespace DotnetPoi.XSSF.UserModel;

/// <summary>
/// XSSF-specific implementation of OOXML data validation.
/// Writes/reads &lt;dataValidation&gt; elements within &lt;dataValidations&gt; in sheet XML.
/// </summary>
public sealed class XSSFDataValidation
{
    /// <summary>Sqref (semi-colon separated range references) like "A1:A10" or "B2:B10 C2:C10".</summary>
    public string Sqref { get; set; } = string.Empty;

    /// <summary>The validation type (whole, decimal, list, etc.). Default None.</summary>
    public DataValidationType Type { get; set; } = DataValidationType.None;

    /// <summary>The validation operator (between, equal, etc.) Default Between.</summary>
    public DataValidationOperator Operator { get; set; } = DataValidationOperator.Between;

    /// <summary>First formula/limit for the validation.</summary>
    public string? Formula1 { get; set; }

    /// <summary>Second formula/limit for the validation (for Between/NotBetween).</summary>
    public string? Formula2 { get; set; }

    /// <summary>Whether blank cells are allowed. Default true.</summary>
    public bool AllowBlank { get; set; } = true;

    /// <summary>Whether to show the input message. Default true.</summary>
    public bool ShowInputMessage { get; set; } = true;

    /// <summary>Whether to show the error alert. Default true.</summary>
    public bool ShowErrorMessage { get; set; } = true;

    /// <summary>Whether to show the drop-down for list type (if false, inline only). Default true.</summary>
    public bool ShowDropDown { get; set; } = true;

    /// <summary>Error alert style: "stop", "warning", "information". Default "stop".</summary>
    public string? ErrorStyle { get; set; }

    /// <summary>Title of the error alert dialog.</summary>
    public string? ErrorTitle { get; set; }

    /// <summary>Error alert message text.</summary>
    public string? ErrorMessage { get; set; }

    /// <summary>Title of the input message prompt.</summary>
    public string? PromptTitle { get; set; }

    /// <summary>Input message text.</summary>
    public string? PromptMessage { get; set; }

    /// <summary>
    /// Returns the first cell range from Sqref, or null if Sqref is empty/invalid.
    /// </summary>
    public CellRangeAddress? GetFirstCellRange()
    {
        if (string.IsNullOrWhiteSpace(Sqref)) return null;
        var parts = Sqref.Split(' ');
        if (parts.Length == 0) return null;
        try { return CellRangeAddress.Parse(parts[0]); }
        catch { return null; }
    }
}
