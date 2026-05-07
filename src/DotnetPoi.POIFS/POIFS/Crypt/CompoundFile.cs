using System.Buffers.Binary;
using System.Text;

namespace DotnetPoi.POIFS.Crypt;

public static class CompoundFile
{
    private static readonly byte[] Signature = { 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1 };

    private const int SectorSize = 512;
    private const int MiniSectorSize = 64;
    private const int DirectoryEntrySize = 128;
    private const int MiniStreamCutoff = 4096;
    private const uint FreeSect = 0xFFFFFFFF;
    private const uint EndOfChain = 0xFFFFFFFE;
    private const uint FatSect = 0xFFFFFFFD;
    private const uint DifSect = 0xFFFFFFFC;
    private const int EndOfChainInt = unchecked((int)EndOfChain);
    private const int NoStream = -1;

    public static void Write(Stream output, IReadOnlyDictionary<string, byte[]> streams)
    {
        Write(output, new CompoundFileDocument(streams));
    }

    public static void Write(Stream output, CompoundFileDocument document)
    {
        Guard.ThrowIfNull(output, nameof(output));
        Guard.ThrowIfNull(document, nameof(document));

        var entries = BuildDirectory(document);
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
        var difatSectors = 0;
        while (true)
        {
            var requiredDifatSectors = fatSectors <= 109 ? 0 : CeilingDiv(fatSectors - 109, 127);
            var requiredFatSectors = CeilingDiv(totalWithoutFat + requiredDifatSectors + fatSectors, SectorSize / sizeof(uint));
            if (requiredFatSectors == fatSectors && requiredDifatSectors == difatSectors)
            {
                break;
            }

            fatSectors = requiredFatSectors;
            difatSectors = requiredDifatSectors;
        }

        var sectorCount = totalWithoutFat + difatSectors + fatSectors;
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

        var difatStart = difatSectors == 0 ? EndOfChainInt : sector;
        for (var i = 0; i < difatSectors; i++)
        {
            fat[difatStart + i] = DifSect;
        }

        sector += difatSectors;

        var fatStart = sector;
        for (var i = 0; i < fatSectors; i++)
        {
            fat[fatStart + i] = FatSect;
        }

        WriteHeader(output, fatSectors, directoryStart, miniFatStart, miniFatSectors, fatStart, difatStart, difatSectors);

        foreach (var entry in regularEntries)
        {
            output.Write(entry.Data, 0, entry.Data.Length);
            WritePadding(output, CeilingDiv(entry.Data.Length, SectorSize) * SectorSize - entry.Data.Length);
        }

        var directoryBytes = WriteDirectory(entries);
        output.Write(directoryBytes, 0, directoryBytes.Length);
        WritePadding(output, directorySectors * SectorSize - directoryBytes.Length);

        WriteFatLikeSectors(output, miniFat, miniFatSectors);
        WriteDifatSectors(output, fatStart, fatSectors, difatStart, difatSectors);
        WriteFatLikeSectors(output, fat, fatSectors);
    }

    public static Dictionary<string, byte[]> ReadStreams(Stream input)
    {
        return ReadStreams(input, preserveStoragePaths: false);
    }

    public static Dictionary<string, byte[]> ReadStreamsWithPaths(Stream input)
    {
        return ReadStreams(input, preserveStoragePaths: true);
    }

    public static CompoundFileDocument ReadDocument(Stream input)
    {
        Guard.ThrowIfNull(input, nameof(input));

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

        var streams = new Dictionary<string, byte[]>(StringComparer.Ordinal);
        var metadata = new Dictionary<string, CompoundFileEntryMetadata>(StringComparer.Ordinal);
        if (root is not null)
        {
            metadata[string.Empty] = root.ToMetadata();
            ReadStorageChildren(
                dirEntries,
                root.ChildId,
                string.Empty,
                entry => ReadEntryStream(file, fat, miniStream, miniFat, miniStreamCutoff, entry),
                streams,
                metadata,
                new HashSet<int>());
        }

        return new CompoundFileDocument(streams, metadata);
    }

    private static Dictionary<string, byte[]> ReadStreams(Stream input, bool preserveStoragePaths)
    {
        Guard.ThrowIfNull(input, nameof(input));

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

        if (preserveStoragePaths && root is not null)
        {
            var pathResults = new Dictionary<string, byte[]>(StringComparer.Ordinal);
            ReadStorageChildren(
                dirEntries,
                root.ChildId,
                string.Empty,
                entry => ReadEntryStream(file, fat, miniStream, miniFat, miniStreamCutoff, entry),
                pathResults,
                new HashSet<int>());
            return pathResults;
        }

        var results = new Dictionary<string, byte[]>(StringComparer.Ordinal);
        foreach (var entry in dirEntries.Where(entry => entry.Type == 2 && !string.IsNullOrEmpty(entry.Name)))
        {
            results[entry.Name] = ReadEntryStream(file, fat, miniStream, miniFat, miniStreamCutoff, entry);
        }

        return results;
    }

    private static List<DirEntry> BuildDirectory(CompoundFileDocument document)
    {
        var root = new DirEntry("Root Entry", 5, Array.Empty<byte>());
        ApplyMetadata(root, string.Empty, document.EntryMetadata);
        foreach (var kv in document.Streams)
        {
            AddStreamEntry(root, kv.Key, kv.Value, document.EntryMetadata);
        }

        var entries = new List<DirEntry>();
        AddEntriesDepthFirst(root, entries);
        for (var i = 0; i < entries.Count; i++)
        {
            entries[i].Index = i;
        }

        foreach (var entry in entries.Where(entry => entry.Type is 1 or 5))
        {
            WirePoiChildTree(entry);
        }

        return entries;
    }

    private static void AddStreamEntry(
        DirEntry root,
        string path,
        byte[] data,
        IReadOnlyDictionary<string, CompoundFileEntryMetadata> metadata)
    {
        var parts = path.Split(new[] { '/' }, StringSplitOptions.RemoveEmptyEntries);
        if (parts.Length == 0)
        {
            return;
        }

        var parent = root;
        var currentPath = string.Empty;
        for (var i = 0; i < parts.Length - 1; i++)
        {
            currentPath = string.IsNullOrEmpty(currentPath) ? parts[i] : currentPath + "/" + parts[i];
            parent = parent.GetOrAddChild(parts[i], 1);
            ApplyMetadata(parent, currentPath, metadata);
        }

        var streamPath = string.IsNullOrEmpty(currentPath)
            ? parts[parts.Length - 1]
            : currentPath + "/" + parts[parts.Length - 1];
        var stream = parent.GetOrAddChild(parts[parts.Length - 1], 2);
        stream.Data = data;
        stream.Size = data.Length;
        ApplyMetadata(stream, streamPath, metadata);
    }

    private static void AddEntriesDepthFirst(DirEntry entry, List<DirEntry> entries)
    {
        entries.Add(entry);
        foreach (var child in entry.Children.OrderBy(child => child.Name, new PoiPropertyNameComparer()))
        {
            AddEntriesDepthFirst(child, entries);
        }
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

            mini.Write(entry.Data, 0, entry.Data.Length);
            WritePadding(mini, miniSectors * MiniSectorSize - entry.Data.Length);
        }

        foreach (var entry in entries.Where(e => e.Type == 2 && !e.IsMiniStream))
        {
            entry.Size = entry.Data.Length;
        }

        miniStream = mini.ToArray();
        return miniFat.ToArray();
    }

    private static void WirePoiChildTree(DirEntry parent)
    {
        var children = parent.Children.OrderBy(e => e.Name, new PoiPropertyNameComparer()).ToArray();
        if (children.Length == 0)
        {
            return;
        }

        var midpoint = children.Length / 2;
        parent.ChildId = children[midpoint].Index;
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

            children[children.Length - 1].LeftSiblingId = NoStream;
            children[children.Length - 1].RightSiblingId = NoStream;
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
            BinaryPrimitives.WriteUInt16LittleEndian(((Span<byte>)buffer).Slice(64), (ushort)Math.Min(nameBytes.Length, 64));
            buffer[66] = entry.Type;
            buffer[67] = entry.Color;
            BinaryPrimitives.WriteInt32LittleEndian(((Span<byte>)buffer).Slice(68), entry.LeftSiblingId);
            BinaryPrimitives.WriteInt32LittleEndian(((Span<byte>)buffer).Slice(72), entry.RightSiblingId);
            BinaryPrimitives.WriteInt32LittleEndian(((Span<byte>)buffer).Slice(76), entry.ChildId);
            entry.ClassId.AsSpan(0, Math.Min(entry.ClassId.Length, 16)).CopyTo(buffer.AsSpan(80, 16));
            BinaryPrimitives.WriteInt32LittleEndian(((Span<byte>)buffer).Slice(96), entry.StateBits);
            BinaryPrimitives.WriteInt64LittleEndian(((Span<byte>)buffer).Slice(100), entry.CreationTime);
            BinaryPrimitives.WriteInt64LittleEndian(((Span<byte>)buffer).Slice(108), entry.ModifiedTime);
            BinaryPrimitives.WriteInt32LittleEndian(((Span<byte>)buffer).Slice(116), entry.StartSector);
            BinaryPrimitives.WriteInt64LittleEndian(((Span<byte>)buffer).Slice(120), entry.Size);
            memory.Write(buffer, 0, buffer.Length);
        }

        return memory.ToArray();
    }

    private static List<DirEntry> ReadDirectory(byte[] directoryBytes)
    {
        var entries = new List<DirEntry>();
        for (var offset = 0; offset + DirectoryEntrySize <= directoryBytes.Length; offset += DirectoryEntrySize)
        {
            var entry = ReadDirectoryEntry(directoryBytes.AsSpan(offset, DirectoryEntrySize));
            entry.Index = entries.Count;
            entries.Add(entry);
        }

        return entries;
    }

    private static DirEntry ReadDirectoryEntry(ReadOnlySpan<byte> buffer)
    {
        var nameLength = BinaryPrimitives.ReadUInt16LittleEndian(buffer.Slice(64));
        var name = string.Empty;
        if (nameLength >= 2)
        {
            name = Encoding.Unicode.GetString(buffer.Slice(0, nameLength - 2).ToArray());
        }

        return new DirEntry(name, buffer[66], Array.Empty<byte>())
        {
            Color = buffer[67],
            LeftSiblingId = BinaryPrimitives.ReadInt32LittleEndian(buffer.Slice(68)),
            RightSiblingId = BinaryPrimitives.ReadInt32LittleEndian(buffer.Slice(72)),
            ChildId = BinaryPrimitives.ReadInt32LittleEndian(buffer.Slice(76)),
            ClassId = buffer.Slice(80, 16).ToArray(),
            StateBits = BinaryPrimitives.ReadInt32LittleEndian(buffer.Slice(96)),
            CreationTime = BinaryPrimitives.ReadInt64LittleEndian(buffer.Slice(100)),
            ModifiedTime = BinaryPrimitives.ReadInt64LittleEndian(buffer.Slice(108)),
            StartSector = BinaryPrimitives.ReadInt32LittleEndian(buffer.Slice(116)),
            Size = BinaryPrimitives.ReadInt64LittleEndian(buffer.Slice(120))
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

        var difatSector = BinaryPrimitives.ReadInt32LittleEndian(file.AsSpan(68));
        var difatSectorCount = BinaryPrimitives.ReadInt32LittleEndian(file.AsSpan(72));
        for (var i = 0; i < difatSectorCount && difatSector >= 0; i++)
        {
            var offset = SectorOffset(difatSector);
            if (offset < 0 || offset + SectorSize > file.Length)
            {
                break;
            }

            for (var j = 0; j < 127 && difat.Count < fatSectorCount; j++)
            {
                var sector = BinaryPrimitives.ReadUInt32LittleEndian(file.AsSpan(offset + j * sizeof(uint)));
                if (sector != FreeSect)
                {
                    difat.Add((int)sector);
                }
            }

            difatSector = BinaryPrimitives.ReadInt32LittleEndian(file.AsSpan(offset + 127 * sizeof(uint)));
        }

        var fat = new List<uint>();
        foreach (var fatSector in difat.Take(fatSectorCount))
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

    private static byte[] ReadEntryStream(
        byte[] file,
        IReadOnlyList<uint> fat,
        byte[] miniStream,
        IReadOnlyList<uint> miniFat,
        uint miniStreamCutoff,
        DirEntry entry)
    {
        return entry.Size < miniStreamCutoff && entry.StartSector >= 0
            ? ReadMiniChain(miniStream, miniFat, entry.StartSector, checked((int)entry.Size))
            : ReadChain(file, fat, entry.StartSector, checked((int)entry.Size));
    }

    private static void ReadStorageChildren(
        IReadOnlyList<DirEntry> entries,
        int childId,
        string prefix,
        Func<DirEntry, byte[]> readStream,
        IDictionary<string, byte[]> results,
        ISet<int> visited)
    {
        ReadStorageChildren(entries, childId, prefix, readStream, results, null, visited);
    }

    private static void ReadStorageChildren(
        IReadOnlyList<DirEntry> entries,
        int childId,
        string prefix,
        Func<DirEntry, byte[]> readStream,
        IDictionary<string, byte[]> results,
        IDictionary<string, CompoundFileEntryMetadata>? metadata,
        ISet<int> visited)
    {
        if (childId < 0 || childId >= entries.Count || !visited.Add(childId))
        {
            return;
        }

        var entry = entries[childId];
        ReadStorageChildren(entries, entry.LeftSiblingId, prefix, readStream, results, metadata, visited);

        if (entry.Type == 1 && !string.IsNullOrEmpty(entry.Name))
        {
            var path = prefix + entry.Name;
            metadata?.Add(path, entry.ToMetadata());
            ReadStorageChildren(entries, entry.ChildId, path + "/", readStream, results, metadata, visited);
        }
        else if (entry.Type == 2 && !string.IsNullOrEmpty(entry.Name))
        {
            var path = prefix + entry.Name;
            results[path] = readStream(entry);
            metadata?.Add(path, entry.ToMetadata());
        }

        ReadStorageChildren(entries, entry.RightSiblingId, prefix, readStream, results, metadata, visited);
    }

    private static void ApplyMetadata(
        DirEntry entry,
        string path,
        IReadOnlyDictionary<string, CompoundFileEntryMetadata> metadata)
    {
        if (!metadata.TryGetValue(path, out var value))
        {
            return;
        }

        entry.Color = value.Color;
        entry.ClassId = value.ClassId.ToArray();
        entry.StateBits = value.StateBits;
        entry.CreationTime = value.CreationTime;
        entry.ModifiedTime = value.ModifiedTime;
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

    private static void WriteHeader(
        Stream output,
        int fatSectors,
        int directoryStart,
        int miniFatStart,
        int miniFatSectors,
        int fatStart,
        int difatStart,
        int difatSectors)
    {
        Span<byte> header = stackalloc byte[SectorSize];
        header.Fill(0xFF);
        Signature.CopyTo(header);
        BinaryPrimitives.WriteUInt16LittleEndian(((Span<byte>)header).Slice(24), 0x003E);
        BinaryPrimitives.WriteUInt16LittleEndian(((Span<byte>)header).Slice(26), 0x0003);
        BinaryPrimitives.WriteUInt16LittleEndian(((Span<byte>)header).Slice(28), 0xFFFE);
        BinaryPrimitives.WriteUInt16LittleEndian(((Span<byte>)header).Slice(30), 9);
        BinaryPrimitives.WriteUInt16LittleEndian(((Span<byte>)header).Slice(32), 6);
        BinaryPrimitives.WriteInt32LittleEndian(((Span<byte>)header).Slice(40), 0);
        BinaryPrimitives.WriteInt32LittleEndian(((Span<byte>)header).Slice(44), fatSectors);
        BinaryPrimitives.WriteInt32LittleEndian(((Span<byte>)header).Slice(48), directoryStart);
        BinaryPrimitives.WriteInt32LittleEndian(((Span<byte>)header).Slice(52), 0);
        BinaryPrimitives.WriteUInt32LittleEndian(((Span<byte>)header).Slice(56), MiniStreamCutoff);
        BinaryPrimitives.WriteInt32LittleEndian(((Span<byte>)header).Slice(60), miniFatStart);
        BinaryPrimitives.WriteInt32LittleEndian(((Span<byte>)header).Slice(64), miniFatSectors);
        BinaryPrimitives.WriteInt32LittleEndian(((Span<byte>)header).Slice(68), difatStart);
        BinaryPrimitives.WriteInt32LittleEndian(((Span<byte>)header).Slice(72), difatSectors);
        for (var i = 0; i < 109; i++)
        {
            BinaryPrimitives.WriteUInt32LittleEndian(((Span<byte>)header).Slice(76 + i * 4), i < fatSectors ? (uint)(fatStart + i) : FreeSect);
        }

        output.Write(header.ToArray(), 0, header.Length);
    }

    private static void WriteDifatSectors(Stream output, int fatStart, int fatSectors, int difatStart, int difatSectors)
    {
        for (var i = 0; i < difatSectors; i++)
        {
            var sector = new byte[SectorSize];
            for (var j = 0; j < 127; j++)
            {
                var fatIndex = 109 + i * 127 + j;
                var value = fatIndex < fatSectors ? (uint)(fatStart + fatIndex) : FreeSect;
                BinaryPrimitives.WriteUInt32LittleEndian(((Span<byte>)sector).Slice(j * sizeof(uint)), value);
            }

            var next = i == difatSectors - 1 ? EndOfChain : (uint)(difatStart + i + 1);
            BinaryPrimitives.WriteUInt32LittleEndian(((Span<byte>)sector).Slice(127 * sizeof(uint)), next);
            output.Write(sector, 0, sector.Length);
        }
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
                BinaryPrimitives.WriteUInt32LittleEndian(((Span<byte>)sector).Slice(j * sizeof(uint)), index < entries.Count ? entries[index] : FreeSect);
            }

            output.Write(sector, 0, sector.Length);
        }
    }

    private static int SectorOffset(int sector) => SectorSize + sector * SectorSize;

    private static int CeilingDiv(int value, int divisor) => value == 0 ? 0 : (value + divisor - 1) / divisor;

    private static void WritePadding(Stream output, int count)
    {
        if (count > 0)
        {
            output.Write(new byte[count], 0, count);
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
        public List<DirEntry> Children { get; } = new();
        public byte Color { get; set; } = 1;
        public byte[] ClassId { get; set; } = new byte[16];
        public int StateBits { get; set; }
        public long CreationTime { get; set; }
        public long ModifiedTime { get; set; }
        public bool IsMiniStream { get; set; }
        public int StartSector { get; set; } = EndOfChainInt;
        public long Size { get; set; }
        public int LeftSiblingId { get; set; } = NoStream;
        public int RightSiblingId { get; set; } = NoStream;
        public int ChildId { get; set; } = NoStream;

        public DirEntry GetOrAddChild(string name, byte type)
        {
            var child = Children.FirstOrDefault(entry =>
                entry.Type == type && string.Equals(entry.Name, name, StringComparison.Ordinal));
            if (child is not null)
            {
                return child;
            }

            child = new DirEntry(name, type, Array.Empty<byte>());
            Children.Add(child);
            return child;
        }

        public CompoundFileEntryMetadata ToMetadata() =>
            new(Type, Color, ClassId.ToArray(), StateBits, CreationTime, ModifiedTime);
    }
}

public sealed class CompoundFileDocument
{
    public CompoundFileDocument(IReadOnlyDictionary<string, byte[]> streams)
        : this(streams, new Dictionary<string, CompoundFileEntryMetadata>(StringComparer.Ordinal))
    {
    }

    public CompoundFileDocument(
        IReadOnlyDictionary<string, byte[]> streams,
        IReadOnlyDictionary<string, CompoundFileEntryMetadata> entryMetadata)
    {
        Streams = streams.ToDictionary(kv => kv.Key, kv => kv.Value, StringComparer.Ordinal);
        EntryMetadata = entryMetadata.ToDictionary(kv => kv.Key, kv => kv.Value, StringComparer.Ordinal);
    }

    public Dictionary<string, byte[]> Streams { get; }

    public Dictionary<string, CompoundFileEntryMetadata> EntryMetadata { get; }
}

public sealed class CompoundFileEntryMetadata
{
    public CompoundFileEntryMetadata(
        byte type,
        byte color,
        byte[] classId,
        int stateBits,
        long creationTime,
        long modifiedTime)
    {
        Type = type;
        Color = color;
        ClassId = classId.ToArray();
        StateBits = stateBits;
        CreationTime = creationTime;
        ModifiedTime = modifiedTime;
    }

    public byte Type { get; }

    public byte Color { get; }

    public byte[] ClassId { get; }

    public int StateBits { get; }

    public long CreationTime { get; }

    public long ModifiedTime { get; }
}
