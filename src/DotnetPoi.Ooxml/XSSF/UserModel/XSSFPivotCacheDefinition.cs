namespace DotnetPoi.XSSF.UserModel;

/// <summary>
/// Represents a pivot cache definition (xl/pivotCache/pivotCacheDefinition{id}.xml).
/// Stores the source data reference (sheet name + cell range).
/// Ported from Apache POI XSSFPivotCacheDefinition.
/// </summary>
public sealed class XSSFPivotCacheDefinition
{
    /// <summary>
    /// The cache ID (matches the cacheId in XSSFPivotCache).
    /// </summary>
    public int CacheId { get; set; }

    /// <summary>
    /// Name of the source sheet that contains the pivot source data.
    /// </summary>
    public string? SourceSheetName { get; set; }

    /// <summary>
    /// Source cell range reference (e.g. "A1:C100").
    /// </summary>
    public string? SourceRef { get; set; }

    /// <summary>
    /// Named range reference (mutually exclusive with SourceRef).
    /// </summary>
    public string? SourceName { get; set; }

    public XSSFPivotCacheDefinition()
    {
    }

    public XSSFPivotCacheDefinition(int cacheId, string sourceSheetName, string sourceRef)
    {
        CacheId = cacheId;
        SourceSheetName = sourceSheetName;
        SourceRef = sourceRef;
    }
}
