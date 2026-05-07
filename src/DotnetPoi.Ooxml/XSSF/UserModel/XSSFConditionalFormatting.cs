using DotnetPoi.SS.UserModel;

namespace DotnetPoi.XSSF.UserModel;

/// <summary>
/// OOXML conditional formatting type attribute values.
/// Only commonly-used types are defined; colorScale/dataBar/iconSet are not yet supported.
/// </summary>
public enum ConditionalFormatType
{
    CellIs,
    Formula,
    Top10,
    UniqueValues,
    DuplicateValues,
    ContainsText,
    NotContainsText,
    BeginsWith,
    EndsWith,
    ContainsBlanks,
    NotContainsBlanks,
    ContainsErrors,
    NotContainsErrors,
    TimePeriod,
    AboveAverage
}

/// <summary>
/// Represents a single conditional formatting rule (&lt;cfRule&gt;).
/// </summary>
public sealed class XSSFCFRule
{
    /// <summary>The rule type (e.g. CellIs, Formula).</summary>
    public ConditionalFormatType Type { get; set; }

    /// <summary>Rule priority (1-based). Default 1.</summary>
    public int Priority { get; set; } = 1;

    /// <summary>Operator for CellIs type: "equal", "notEqual", "greaterThan", "lessThan", etc.</summary>
    public string? Operator { get; set; }

    /// <summary>Text value for text-based rules (containsText, beginsWith, etc.).</summary>
    public string? Text { get; set; }

    /// <summary>Index into the differential formatting table (&lt;dxfs&gt;). -1 means not set.</summary>
    public int DxfId { get; set; } = -1;

    /// <summary>List of formula strings.</summary>
    public List<string> Formulas { get; } = new();
}

/// <summary>
/// XSSF-specific implementation of OOXML conditional formatting.
/// Writes/reads &lt;conditionalFormatting&gt; elements in sheet XML.
/// </summary>
public sealed class XSSFConditionalFormatting
{
    /// <summary>Sqref range(s) the formatting applies to (e.g. "A1:A10").</summary>
    public string Sqref { get; set; } = string.Empty;

    /// <summary>List of rules.</summary>
    public List<XSSFCFRule> Rules { get; } = new();
}
