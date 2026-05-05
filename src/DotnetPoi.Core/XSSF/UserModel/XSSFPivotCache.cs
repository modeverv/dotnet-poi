namespace DotnetPoi.XSSF.UserModel;

/// <summary>
/// Represents a pivot cache reference stored in the workbook's pivotCaches element.
/// Ported from Apache POI XSSFPivotCache.
/// </summary>
public sealed class XSSFPivotCache
{
    /// <summary>
    /// The cache ID (0-based, assigned sequentially as pivot tables are created).
    /// </summary>
    public int CacheId { get; set; }

    public XSSFPivotCache()
    {
    }

    public XSSFPivotCache(int cacheId)
    {
        CacheId = cacheId;
    }
}
