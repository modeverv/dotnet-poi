using System.Buffers.Binary;
using System.Text;

namespace DotnetPoi.POIFS.Crypt;

internal static class CompoundFile
{
    private static readonly byte[] Signature = { 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1 };

    private const int SectorSize = 512;
    private const int MiniSectorSize = 64;
    private const int DirectoryEntrySize = 128;
    private const int MiniStreamCutoff = 4096;
    private const uint FreeSect = 0xFFFFFFFF;
    private const uint EndOfChain = 0xFFFFFFFE;
    private const uint FatSect = 0xFFFFFFFD;
    private const int EndOfChainInt = unchecked((int)EndOfChain);
    private const int NoStream = -1;

    public static void Write(Stream output, IReadOnlyDictionary<string, byte[]> streams)
    {
        ArgumentNullException.ThrowIfNull(output);
        ArgumentNullException.ThrowIfNull(streams);

        var entries = BuildDirectory(streams);
        var miniFat = BuildMiniStreams(entries, out var miniStream);
        entries[0].Data = miniStream;
        entries[0].Size = miniStream.Length;

        var regularEntries = entries
            .Where(e => e.Type == 2 && !e.IsMiniStream && e.Data.Length > 0)
            .Concat(entries[0].Data.Length > 0 ? new[] { entries[0] } : Array.Empty<DirEntry>())
            .ToList();

        var directorySectors = CeilingDiv(entries.Count * DirectoryEntrySize, SectorSize);
        var miniFatSectors = CeilingDiv(miniFat.Length * sizeof(uint), SectorSize);
        var regularStreamSectors = regularEntries.Sum(e => CeilingDiv(e.Data.Length, SectorSize));
        var totalWithoutFat = regularStreamSectors + directorySectors + miniFatSectors;
        var fatSectors = 1;
        while (CeilingDiv(totalWithoutFat + fatSectors, SectorSize / sizeof(uint)) != fatSectors)
        {
            fatSectors = CeilingDiv(totalWithoutFat + fatSectors, SectorSize / sizeof(uint));
        }

        var sectorCount = totalWithoutFat + fatSectors;
        var fat = Enumerable.Repeat(FreeSect, sectorCount).ToArray();
        var sector = 0;
        foreach (var entry in regularEntries)
        {
            entry.StartSector = sector;
            var sectorCountForEntry = CeilingDiv(entry.Data.Length, SectorSize);
            MarkChain(fat, sector, sectorCountForEntry);
            sector += sectorCountForEntry;
        }

        var directoryStart = sector;
        MarkChain(fat, directoryStart, directorySectors);
        sector += directorySectors;

        var miniFatStart = miniFatSectors == 0 ? EndOfChainInt : sector;
        if (miniFatSectors > 0)
        {
            MarkChain(fat, miniFatStart, miniFatSectors);
            sector += miniFatSectors;
        }

        var fatStart = sector;
        for (var i = 0; i < fatSectors; i++)
        {
            fat[fatStart + i] = FatSect;
        }

        WriteHeader(output, fatSectors, directoryStart, miniFatStart, miniFatSectors, fatStart);

        foreach (var entry in regularEntries)
        {
            output.Write(entry.Data);
            WritePadding(output, CeilingDiv(entry.Data.Length, SectorSize) * SectorSize - entry.Data.Length);
        }

        var directoryBytes = WriteDirectory(entries);
        output.Write(directoryBytes);
        WritePadding(output, directorySectors * SectorSize - directoryBytes.Length);

        WriteFatLikeSectors(output, miniFat, miniFatSectors);
        WriteFatLikeSectors(output, fat, fatSectors);
    }

    public static Dictionary<string, byte[]> ReadStreams(Stream input)
    {
        ArgumentNullException.ThrowIfNull(input);

        using var memory = new MemoryStream();
        input.CopyTo(memory);
        var file = memory.ToArray();
        if (file.Length < SectorSize || !file.AsSpan(0, Signature.Length).SequenceEqual(Signature))
        {
            throw new InvalidDataException("Not an OLE2 compound document.");
        }

        var fatSectorCount = BinaryPrimitives.ReadInt32LittleEndian(file.AsSpan(44));
        var directoryStart = BinaryPrimitives.ReadInt32LittleEndian(file.AsSpan(48));
        var miniStreamCutoff = BinaryPrimitives.ReadUInt32LittleEndian(file.AsSpan(56));
        var miniFatStart = BinaryPrimitives.ReadInt32LittleEndian(file.AsSpan(60));
        var miniFatSectorCount = BinaryPrimitives.ReadInt32LittleEndian(file.AsSpan(64));
        var fat = ReadFat(file, fatSectorCount);

        var directoryBytes = ReadChain(file, fat, directoryStart, int.MaxValue);
        var dirEntries = ReadDirectory(directoryBytes);
        var root = dirEntries.FirstOrDefault(e => e.Type == 5);
        var miniStream = root is not null && root.StartSector >= 0 && root.Size > 0
            ? ReadChain(file, fat, root.StartSector, checked((int)root.Size))
            : Array.Empty<byte>();
        var miniFat = miniFatStart >= 0 && miniFatSectorCount > 0
            ? ReadFatEntries(ReadChain(file, fat, miniFatStart, checked(miniFatSectorCount * SectorSize)))
            : Array.Empty<uint>();

        var results = new Dictionary<string, byte[]>(StringComparer.Ordinal);
        foreach (var entry in dirEntries)
        {
            if (entry.Type != 2 || string.IsNullOrEmpty(entry.Name))
            {
                continue;
            }

            results[entry.Name] = entry.Size < miniStreamCutoff && entry.StartSector >= 0
                ? ReadMiniChain(miniStream, miniFat, entry.StartSector, checked((int)entry.Size))
                : ReadChain(file, fat, entry.StartSector, checked((int)entry.Size));
        }

        return results;
    }

    private static List<DirEntry> BuildDirectory(IReadOnlyDictionary<string, byte[]> streams)
    {
        var entries = new List<DirEntry> { new("Root Entry", 5, Array.Empty<byte>()) };
        foreach (var (name, data) in streams.OrderBy(kv => kv.Key, new PoiPropertyNameComparer()))
        {
            entries.Add(new DirEntry(name, 2, data));
        }

        for (var i = 0; i < entries.Count; i++)
        {
            entries[i].Index = i;
        }

        WirePoiChildTree(entries);
        return entries;
    }

    private static uint[] BuildMiniStreams(List<DirEntry> entries, out byte[] miniStream)
    {
        var miniFat = new List<uint>();
        using var mini = new MemoryStream();
        foreach (var entry in entries.Where(e => e.Type == 2 && e.Data.Length < MiniStreamCutoff))
        {
            entry.IsMiniStream = true;
            entry.Size = entry.Data.Length;
            if (entry.Data.Length == 0)
            {
                entry.StartSector = EndOfChainInt;
                continue;
            }

            entry.StartSector = miniFat.Count;
            var miniSectors = CeilingDiv(entry.Data.Length, MiniSectorSize);
            for (var i = 0; i < miniSectors; i++)
            {
                miniFat.Add(i == miniSectors - 1 ? EndOfChain : (uint)(entry.StartSector + i + 1));
            }

            mini.Write(entry.Data);
            WritePadding(mini, miniSectors * MiniSectorSize - entry.Data.Length);
        }

        foreach (var entry in entries.Where(e => e.Type == 2 && !e.IsMiniStream))
        {
            entry.Size = entry.Data.Length;
        }

        miniStream = mini.ToArray();
        return miniFat.ToArray();
    }

    private static void WirePoiChildTree(List<DirEntry> entries)
    {
        var children = entries.Skip(1).OrderBy(e => e.Name, new PoiPropertyNameComparer()).ToArray();
        if (children.Length == 0)
        {
            return;
        }

        var midpoint = children.Length / 2;
        entries[0].ChildId = children[midpoint].Index;
        children[0].LeftSiblingId = NoStream;
        children[0].RightSiblingId = NoStream;
        for (var j = 1; j < midpoint; j++)
        {
            children[j].LeftSiblingId = children[j - 1].Index;
            children[j].RightSiblingId = NoStream;
        }

        if (midpoint != 0)
        {
            children[midpoint].LeftSiblingId = children[midpoint - 1].Index;
        }

        if (midpoint != children.Length - 1)
        {
            children[midpoint].RightSiblingId = children[midpoint + 1].Index;
            for (var j = midpoint + 1; j < children.Length - 1; j++)
            {
                children[j].LeftSiblingId = NoStream;
                children[j].RightSiblingId = children[j + 1].Index;
            }

            children[^1].LeftSiblingId = NoStream;
            children[^1].RightSiblingId = NoStream;
        }
        else
        {
            children[midpoint].RightSiblingId = NoStream;
        }
    }

    private static byte[] WriteDirectory(List<DirEntry> entries)
    {
        using var memory = new MemoryStream();
        foreach (var entry in entries)
        {
            var buffer = new byte[DirectoryEntrySize];
            var nameBytes = Encoding.Unicode.GetBytes(entry.Name + "\0");
            nameBytes.AsSpan(0, Math.Min(nameBytes.Length, 64)).CopyTo(buffer);
            BinaryPrimitives.WriteUInt16LittleEndian(buffer.AsSpan(64), (ushort)Math.Min(nameBytes.Length, 64));
            buffer[66] = entry.Type;
            buffer[67] = 1;
            BinaryPrimitives.WriteInt32LittleEndian(buffer.AsSpan(68), entry.LeftSiblingId);
            BinaryPrimitives.WriteInt32LittleEndian(buffer.AsSpan(72), entry.RightSiblingId);
            BinaryPrimitives.WriteInt32LittleEndian(buffer.AsSpan(76), entry.ChildId);
            BinaryPrimitives.WriteInt32LittleEndian(buffer.AsSpan(116), entry.StartSector);
            BinaryPrimitives.WriteInt64LittleEndian(buffer.AsSpan(120), entry.Size);
            memory.Write(buffer);
        }

        return memory.ToArray();
    }

    private static List<DirEntry> ReadDirectory(byte[] directoryBytes)
    {
        var entries = new List<DirEntry>();
        for (var offset = 0; offset + DirectoryEntrySize <= directoryBytes.Length; offset += DirectoryEntrySize)
        {
            var entry = ReadDirectoryEntry(directoryBytes.AsSpan(offset, DirectoryEntrySize));
            if (entry.Type == 0 && string.IsNullOrEmpty(entry.Name))
            {
                continue;
            }

            entries.Add(entry);
        }

        return entries;
    }

    private static DirEntry ReadDirectoryEntry(ReadOnlySpan<byte> buffer)
    {
        var nameLength = BinaryPrimitives.ReadUInt16LittleEndian(buffer[64..]);
        var name = string.Empty;
        if (nameLength >= 2)
        {
            name = Encoding.Unicode.GetString(buffer[..(nameLength - 2)]);
        }

        return new DirEntry(name, buffer[66], Array.Empty<byte>())
        {
            LeftSiblingId = BinaryPrimitives.ReadInt32LittleEndian(buffer[68..]),
            RightSiblingId = BinaryPrimitives.ReadInt32LittleEndian(buffer[72..]),
            ChildId = BinaryPrimitives.ReadInt32LittleEndian(buffer[76..]),
            StartSector = BinaryPrimitives.ReadInt32LittleEndian(buffer[116..]),
            Size = BinaryPrimitives.ReadInt64LittleEndian(buffer[120..])
        };
    }

    private static List<uint> ReadFat(byte[] file, int fatSectorCount)
    {
        var difat = new List<int>();
        for (var i = 0; i < Math.Min(fatSectorCount, 109); i++)
        {
            var sector = BinaryPrimitives.ReadUInt32LittleEndian(file.AsSpan(76 + i * 4));
            if (sector != FreeSect)
            {
                difat.Add((int)sector);
            }
        }

        var fat = new List<uint>();
        foreach (var fatSector in difat)
        {
            var offset = SectorOffset(fatSector);
            for (var i = 0; i < SectorSize / sizeof(uint); i++)
            {
                fat.Add(BinaryPrimitives.ReadUInt32LittleEndian(file.AsSpan(offset + i * 4)));
            }
        }

        return fat;
    }

    private static uint[] ReadFatEntries(byte[] bytes)
    {
        var fat = new uint[bytes.Length / sizeof(uint)];
        for (var i = 0; i < fat.Length; i++)
        {
            fat[i] = BinaryPrimitives.ReadUInt32LittleEndian(bytes.AsSpan(i * 4));
        }

        return fat;
    }

    private static byte[] ReadChain(byte[] file, IReadOnlyList<uint> fat, int startSector, int length)
    {
        if (startSector < 0)
        {
            return Array.Empty<byte>();
        }

        using var output = new MemoryStream();
        var sector = startSector;
        while (sector >= 0 && sector < fat.Count)
        {
            var offset = SectorOffset(sector);
            var copy = Math.Min(SectorSize, file.Length - offset);
            if (copy <= 0)
            {
                break;
            }

            output.Write(file, offset, copy);
            if (fat[sector] == EndOfChain)
            {
                break;
            }

            sector = (int)fat[sector];
        }

        return Truncate(output.ToArray(), length);
    }

    private static byte[] ReadMiniChain(byte[] miniStream, IReadOnlyList<uint> miniFat, int startSector, int length)
    {
        if (startSector < 0)
        {
            return Array.Empty<byte>();
        }

        using var output = new MemoryStream();
        var sector = startSector;
        while (sector >= 0 && sector < miniFat.Count)
        {
            var offset = sector * MiniSectorSize;
            var copy = Math.Min(MiniSectorSize, miniStream.Length - offset);
            if (copy <= 0)
            {
                break;
            }

            output.Write(miniStream, offset, copy);
            if (miniFat[sector] == EndOfChain)
            {
                break;
            }

            sector = (int)miniFat[sector];
        }

        return Truncate(output.ToArray(), length);
    }

    private static byte[] Truncate(byte[] bytes, int length)
    {
        if (length != int.MaxValue && bytes.Length > length)
        {
            Array.Resize(ref bytes, length);
        }

        return bytes;
    }

    private static void WriteHeader(Stream output, int fatSectors, int directoryStart, int miniFatStart, int miniFatSectors, int fatStart)
    {
        Span<byte> header = stackalloc byte[SectorSize];
        header.Fill(0xFF);
        Signature.CopyTo(header);
        BinaryPrimitives.WriteUInt16LittleEndian(header[24..], 0x003E);
        BinaryPrimitives.WriteUInt16LittleEndian(header[26..], 0x0003);
        BinaryPrimitives.WriteUInt16LittleEndian(header[28..], 0xFFFE);
        BinaryPrimitives.WriteUInt16LittleEndian(header[30..], 9);
        BinaryPrimitives.WriteUInt16LittleEndian(header[32..], 6);
        BinaryPrimitives.WriteInt32LittleEndian(header[40..], 0);
        BinaryPrimitives.WriteInt32LittleEndian(header[44..], fatSectors);
        BinaryPrimitives.WriteInt32LittleEndian(header[48..], directoryStart);
        BinaryPrimitives.WriteInt32LittleEndian(header[52..], 0);
        BinaryPrimitives.WriteUInt32LittleEndian(header[56..], MiniStreamCutoff);
        BinaryPrimitives.WriteInt32LittleEndian(header[60..], miniFatStart);
        BinaryPrimitives.WriteInt32LittleEndian(header[64..], miniFatSectors);
        BinaryPrimitives.WriteUInt32LittleEndian(header[68..], EndOfChain);
        BinaryPrimitives.WriteInt32LittleEndian(header[72..], 0);
        for (var i = 0; i < 109; i++)
        {
            BinaryPrimitives.WriteUInt32LittleEndian(header[(76 + i * 4)..], i < fatSectors ? (uint)(fatStart + i) : FreeSect);
        }

        output.Write(header);
    }

    private static void MarkChain(uint[] fat, int startSector, int sectorCount)
    {
        for (var i = 0; i < sectorCount; i++)
        {
            fat[startSector + i] = i == sectorCount - 1 ? EndOfChain : (uint)(startSector + i + 1);
        }
    }

    private static void WriteFatLikeSectors(Stream output, IReadOnlyList<uint> entries, int sectorCount)
    {
        for (var i = 0; i < sectorCount; i++)
        {
            var sector = new byte[SectorSize];
            for (var j = 0; j < SectorSize / sizeof(uint); j++)
            {
                var index = i * (SectorSize / sizeof(uint)) + j;
                BinaryPrimitives.WriteUInt32LittleEndian(sector.AsSpan(j * sizeof(uint)), index < entries.Count ? entries[index] : FreeSect);
            }

            output.Write(sector);
        }
    }

    private static int SectorOffset(int sector) => SectorSize + sector * SectorSize;

    private static int CeilingDiv(int value, int divisor) => value == 0 ? 0 : (value + divisor - 1) / divisor;

    private static void WritePadding(Stream output, int count)
    {
        if (count > 0)
        {
            output.Write(new byte[count]);
        }
    }

    private sealed class PoiPropertyNameComparer : IComparer<string>
    {
        public int Compare(string? x, string? y)
        {
            x ??= string.Empty;
            y ??= string.Empty;
            var result = x.Length - y.Length;
            return result != 0 ? result : string.Compare(x, y, StringComparison.OrdinalIgnoreCase);
        }
    }

    private sealed class DirEntry
    {
        public DirEntry(string name, byte type, byte[] data)
        {
            Name = name;
            Type = type;
            Data = data;
            Size = data.Length;
        }

        public int Index { get; set; }
        public string Name { get; }
        public byte Type { get; }
        public byte[] Data { get; set; }
        public bool IsMiniStream { get; set; }
        public int StartSector { get; set; } = EndOfChainInt;
        public long Size { get; set; }
        public int LeftSiblingId { get; set; } = NoStream;
        public int RightSiblingId { get; set; } = NoStream;
        public int ChildId { get; set; } = NoStream;
    }
}
