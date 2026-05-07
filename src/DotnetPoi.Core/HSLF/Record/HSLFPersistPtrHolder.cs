using System.Buffers.Binary;
using System.Collections.Generic;

namespace DotnetPoi.HSLF.Record;

/// <summary>
/// Holds persist pointer mappings from PersistPtrIncrementalBlock (6002)
/// and PersistPtrFullBlock (6001) records.
///
/// These records map persist IDs to byte offsets in the PowerPoint Document stream,
/// linking SlidePersistAtom refIDs to the actual Slide (1006) record locations.
///
/// Ported from org.apache.poi.hslf.record.PersistPtrHolder.
/// [MS-PPT] § 2.3.7 (PersistPtrIncrementalBlock) / § 2.3.6 (PersistPtrFullBlock)
/// </summary>
internal sealed class HSLFPersistPtrHolder
{
    /// <summary>Maps persist ID → byte offset in the PowerPoint Document stream.</summary>
    public IReadOnlyDictionary<int, int> SlideLocations { get; }

    private HSLFPersistPtrHolder(Dictionary<int, int> locations)
    {
        SlideLocations = locations;
    }

    /// <summary>
    /// Parses all PersistPtrHolder records (6001 and 6002) from top-level records
    /// and merges their mappings into a single persistId → offset dictionary.
    ///
    /// Later records (by offset) override earlier ones for the same persist ID.
    /// </summary>
    public static Dictionary<int, int> BuildPersistMap(IReadOnlyList<HSLFRecord> topLevelRecords)
    {
        // Collect persist ptr records and parse their maps
        var combined = new Dictionary<int, int>();

        foreach (var record in topLevelRecords)
        {
            if (record.RecType != 6001 && record.RecType != 6002)
                continue;
            if (!record.IsAtom)
                continue;

            var map = ParseSingle(record);
            foreach (var kv in map)
            {
                // Later records override earlier ones (most recent wins)
                combined[kv.Key] = kv.Value;
            }
        }

        return combined;
    }

    /// <summary>
    /// Parses a single PersistPtrFullBlock (6001) or PersistPtrIncrementalBlock (6002) atom.
    ///
    /// Body format:
    ///   Repeated groups of:
    ///     4 bytes: info (20 bits offset_no + 12 bits offset_count)
    ///     offset_count × 4 bytes: sheet offsets (unsigned 32-bit)
    ///
    /// Each entry maps persist ID = offset_no + i to the byte offset where the
    /// Slide/Notes/Document record lives in the PowerPoint Document stream.
    /// </summary>
    internal static Dictionary<int, int> ParseSingle(HSLFRecord record)
    {
        var result = new Dictionary<int, int>();
        if (record.RecordLength < 4)
            return result;

        var body = record.Body;
        int pos = 0;

        while (pos + 4 <= body.Length)
        {
            int info = BinaryPrimitives.ReadInt32LittleEndian(body.Slice(pos));

            // First 20 bits: persist ID base
            int offsetNo = info & 0x000FFFFF;
            // Remaining 12 bits: entry count
            int offsetCount = (info >> 20) & 0xFFF;

            pos += 4;

            for (int i = 0; i < offsetCount && pos + 4 <= body.Length; i++)
            {
                int sheetOffset = (int)BinaryPrimitives.ReadUInt32LittleEndian(body.Slice(pos));
                int persistId = offsetNo + i;
                result[persistId] = sheetOffset;
                pos += 4;
            }
        }

        return result;
    }
}
