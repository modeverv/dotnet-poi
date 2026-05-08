using System.Buffers.Binary;

namespace DotnetPoi.HSLF.Record;

/// <summary>
/// Represents a single record in the PowerPoint binary record tree.
/// Ported from org.apache.poi.hslf.record.Record + RecordContainer + RecordAtom.
///
/// Each record has an 8-byte header (verAndInstance + recType + recLen),
/// followed by body data (atoms) or child records (containers).
/// </summary>
public abstract class HSLFRecord
{
    protected HSLFRecord(
        ushort recType,
        ushort verAndInstance,
        int recLen,
        int offset,
        byte[] rawBytes)
    {
        RecType = recType;
        VerAndInstance = verAndInstance;
        RecordLength = recLen;
        Offset = offset;
        RawBytes = rawBytes;
    }

    /// <summary>The record type identifier (e.g. 1006 for Slide, 4000 for TextCharsAtom).</summary>
    public ushort RecType { get; }

    /// <summary>
    /// The combined version (high nibble) and instance (low nibble) field.
    /// For containers, the low nibble is 0x0F.
    /// </summary>
    public ushort VerAndInstance { get; }

    /// <summary>The length of the record body (not including the 8-byte header).</summary>
    public int RecordLength { get; }

    /// <summary>Byte offset of this record's header within the PowerPoint Document stream.</summary>
    public int Offset { get; }

    /// <summary>
    /// Full raw bytes of this record, including the 8-byte header and all body/child data.
    /// Length = RawBytesHeaderSize + RecordLength.
    /// </summary>
    public byte[] RawBytes { get; }

    /// <summary>Returns the child records if this is a container, or null if this is an atom.</summary>
    public abstract IReadOnlyList<HSLFRecord>? Children { get; }

    /// <summary>True when this record is a container (can have child records).</summary>
    public bool IsContainer => (VerAndInstance & 0x0F) == 0x0F;

    /// <summary>True when this record is an atom (leaf node, no child records).</summary>
    public bool IsAtom => !IsContainer;

    /// <summary>Size of the record header in bytes.</summary>
    public const int RawBytesHeaderSize = 8;

    /// <summary>Returns the raw body bytes (everything after the 8-byte header).</summary>
    public ReadOnlySpan<byte> Body => RawBytes.AsSpan(RawBytesHeaderSize, RecordLength);
}

/// <summary>
/// A container record that holds child records.
/// Ported from org.apache.poi.hslf.record.RecordContainer.
/// </summary>
public sealed class HSLFRecordContainer : HSLFRecord
{
    private readonly IReadOnlyList<HSLFRecord> _children;

    internal HSLFRecordContainer(
        ushort recType,
        ushort verAndInstance,
        int recLen,
        int offset,
        byte[] rawBytes,
        IReadOnlyList<HSLFRecord> children)
        : base(recType, verAndInstance, recLen, offset, rawBytes)
    {
        _children = children;
    }

    /// <summary>Child records of this container. May be empty but never null.</summary>
    public override IReadOnlyList<HSLFRecord>? Children => _children;
}

/// <summary>
/// An atom record that holds data (no child records).
/// Ported from org.apache.poi.hslf.record.RecordAtom.
/// </summary>
public sealed class HSLFRecordAtom : HSLFRecord
{
    internal HSLFRecordAtom(
        ushort recType,
        ushort verAndInstance,
        int recLen,
        int offset,
        byte[] rawBytes)
        : base(recType, verAndInstance, recLen, offset, rawBytes)
    {
    }

    /// <summary>Atoms have no children.</summary>
    public override IReadOnlyList<HSLFRecord>? Children => null;
}
