using System.Buffers.Binary;
using System.Text;

namespace DotnetPoi.POIFS.Crypt;

internal static class CompoundFile
{
    private const int SectorSize = 512;
    private const int DirectoryEntrySize = 128;
    private const uint FreeSect = 0xFFFFFFFF;
    private const uint EndOfChain = 0xFFFFFFFE;
    private const uint FatSect = 0xFFFFFFFD;
    private const int NoStream = -1;

    public static void Write(Stream output, IReadOnlyDictionary<string, byte[]> streams)
    {
        var entries = BuildDirectory(streams);
        var directorySectors = CeilingDiv(entries.Count * DirectoryEntrySize, SectorSize);
        var streamSectorCount = entries.Where(e => e.Type == 2).Sum(e => e.SectorCount);
        var totalWithoutFat = directorySectors + streamSectorCount;
        var fatSectors = 1;
        while (CeilingDiv(totalWithoutFat + fatSectors, 128) != fatSectors)
        {
            fatSectors = CeilingDiv(totalWithoutFat + fatSectors, 128);
        }

        var sectorCount = totalWithoutFat + fatSectors;
        var fat = Enumerable.Repeat(FreeSect, sectorCount).ToArray();
        var sector = 0;
        foreach (var entry in entries.Where(e => e.Type == 2))
        {
            entry.StartSector = sector;
            for (var i = 0; i < entry.SectorCount; i++)
            {
                fat[sector + i] = i == entry.SectorCount - 1 ? EndOfChain : (uint)(sector + i + 1);
            }
            sector += entry.SectorCount;
        }

        var directoryStart = sector;
        for (var i = 0; i < directorySectors; i++)
        {
            fat[sector + i] = i == directorySectors - 1 ? EndOfChain : (uint)(sector + i + 1);
        }
        sector += directorySectors;
        var fatStart = sector;
        for (var i = 0; i < fatSectors; i++)
        {
            fat[fatStart + i] = FatSect;
        }

        WriteHeader(output, fatSectors, directoryStart, fatStart);
        var directoryBytes = WriteDirectory(entries);
        foreach (var entry in entries.Where(e => e.Type == 2))
        {
            output.Write(entry.Data);
            WritePadding(output, entry.SectorCount * SectorSize - entry.Data.Length);
        }

        output.Write(directoryBytes);
        WritePadding(output, directorySectors * SectorSize - directoryBytes.Length);

        for (var i = 0; i < fatSectors; i++)
        {
            var fatSector = new byte[SectorSize];
            for (var j = 0; j < 128; j++)
            {
                var index = i * 128 + j;
                BinaryPrimitives.WriteUInt32LittleEndian(fatSector.AsSpan(j * 4), index < fat.Length ? fat[index] : FreeSect);
            }
            output.Write(fatSector);
        }
    }

    public static Dictionary<string, byte[]> ReadStreams(Stream input)
    {
        using var memory = new MemoryStream();
        input.CopyTo(memory);
        var file = memory.ToArray();
        if (file.Length < SectorSize || !file.AsSpan(0, 8).SequenceEqual(new byte[] { 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1 }))
        {
            throw new InvalidDataException("Not an OLE2 compound document.");
        }

        var fatSectorCount = BinaryPrimitives.ReadInt32LittleEndian(file.AsSpan(44));
        var directoryStart = BinaryPrimitives.ReadInt32LittleEndian(file.AsSpan(48));
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
            for (var i = 0; i < 128; i++)
            {
                fat.Add(BinaryPrimitives.ReadUInt32LittleEndian(file.AsSpan(offset + i * 4)));
            }
        }

        var directoryBytes = ReadChain(file, fat, directoryStart, int.MaxValue);
        var dirEntries = new List<DirEntry>();
        for (var offset = 0; offset + DirectoryEntrySize <= directoryBytes.Length; offset += DirectoryEntrySize)
        {
            var entry = ReadDirectoryEntry(directoryBytes.AsSpan(offset, DirectoryEntrySize));
            dirEntries.Add(entry);
        }

        var results = new Dictionary<string, byte[]>(StringComparer.Ordinal);
        for (var i = 0; i < dirEntries.Count; i++)
        {
            var entry = dirEntries[i];
            if (entry.Type != 2 || string.IsNullOrEmpty(entry.Name) || entry.StartSector < 0)
            {
                continue;
            }

            results[entry.Name] = ReadChain(file, fat, entry.StartSector, checked((int)entry.Size));
        }

        return results;
    }

    private static List<DirEntry> BuildDirectory(IReadOnlyDictionary<string, byte[]> streams)
    {
        var entries = new List<DirEntry>
        {
            new("Root Entry", 5, Array.Empty<byte>())
        };

        foreach (var (name, data) in streams.OrderBy(kv => kv.Key, StringComparer.Ordinal))
        {
            entries.Add(new DirEntry(name, 2, data));
        }

        if (entries.Count > 1)
        {
            entries[0].ChildId = 1;
            for (var i = 1; i < entries.Count; i++)
            {
                entries[i].RightSiblingId = i == entries.Count - 1 ? NoStream : i + 1;
            }
        }

        return entries;
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

    private static byte[] ReadChain(byte[] file, IReadOnlyList<uint> fat, int startSector, int length)
    {
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

        var bytes = output.ToArray();
        if (length != int.MaxValue && bytes.Length > length)
        {
            Array.Resize(ref bytes, length);
        }

        return bytes;
    }

    private static void WriteHeader(Stream output, int fatSectors, int directoryStart, int fatStart)
    {
        Span<byte> header = stackalloc byte[SectorSize];
        new byte[] { 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1 }.CopyTo(header);
        BinaryPrimitives.WriteUInt16LittleEndian(header[24..], 0x003E);
        BinaryPrimitives.WriteUInt16LittleEndian(header[26..], 0x0003);
        BinaryPrimitives.WriteUInt16LittleEndian(header[28..], 0xFFFE);
        BinaryPrimitives.WriteUInt16LittleEndian(header[30..], 9);
        BinaryPrimitives.WriteUInt16LittleEndian(header[32..], 6);
        BinaryPrimitives.WriteInt32LittleEndian(header[40..], 0);
        BinaryPrimitives.WriteInt32LittleEndian(header[44..], fatSectors);
        BinaryPrimitives.WriteInt32LittleEndian(header[48..], directoryStart);
        BinaryPrimitives.WriteUInt32LittleEndian(header[56..], 4096);
        BinaryPrimitives.WriteUInt32LittleEndian(header[60..], EndOfChain);
        BinaryPrimitives.WriteInt32LittleEndian(header[64..], 0);
        BinaryPrimitives.WriteUInt32LittleEndian(header[68..], EndOfChain);
        BinaryPrimitives.WriteInt32LittleEndian(header[72..], 0);
        for (var i = 0; i < 109; i++)
        {
            BinaryPrimitives.WriteUInt32LittleEndian(header[(76 + i * 4)..], i < fatSectors ? (uint)(fatStart + i) : FreeSect);
        }

        output.Write(header);
    }

    private static int SectorOffset(int sector) => SectorSize + sector * SectorSize;

    private static int CeilingDiv(int value, int divisor) => (value + divisor - 1) / divisor;

    private static void WritePadding(Stream output, int count)
    {
        if (count <= 0) return;
        output.Write(new byte[count]);
    }

    private sealed class DirEntry
    {
        public DirEntry(string name, byte type, byte[] data)
        {
            Name = name;
            Type = type;
            Data = data;
            Size = data.Length;
            SectorCount = data.Length == 0 ? 0 : CeilingDiv(data.Length, SectorSize);
            StartSector = SectorCount == 0 ? -2 : 0;
        }

        public string Name { get; }
        public byte Type { get; }
        public byte[] Data { get; }
        public long Size { get; set; }
        public int SectorCount { get; }
        public int StartSector { get; set; }
        public int LeftSiblingId { get; set; } = NoStream;
        public int RightSiblingId { get; set; } = NoStream;
        public int ChildId { get; set; } = NoStream;
    }
}
