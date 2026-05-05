namespace DotnetPoi.XSSF.UserModel;

/// <summary>
/// Represents a pivot cache records part (xl/pivotCache/pivotCacheRecords{id}.xml).
/// In the minimal implementation, this is emitted as an empty records container.
/// Ported from Apache POI XSSFPivotCacheRecords.
/// </summary>
public sealed class XSSFPivotCacheRecords
{
    /// <summary>
    /// The cache ID that these records belong to.
    /// </summary>
    public int CacheId { get; set; }

    public XSSFPivotCacheRecords()
    {
    }

    public XSSFPivotCacheRecords(int cacheId)
    {
        CacheId = cacheId;
    }
}
