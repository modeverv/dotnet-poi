using DotnetPoi.SS.UserModel;

namespace DotnetPoi.HSSF.UserModel;

public sealed class HSSFDataFormat : IDataFormat
{
    private readonly Dictionary<string, short> _formatsByString = new(StringComparer.Ordinal);
    private readonly Dictionary<short, string> _formatsByIndex = new();
    private short _nextCustomIndex = 164;

    public HSSFDataFormat()
    {
        _formatsByIndex[0] = "General";
        _formatsByString["General"] = 0;
    }

    public short getFormat(string format)
    {
        Guard.ThrowIfNull(format, nameof(format));
        if (_formatsByString.TryGetValue(format, out var existing))
        {
            return existing;
        }

        var index = _nextCustomIndex++;
        _formatsByString[format] = index;
        _formatsByIndex[index] = format;
        return index;
    }

    public string? getFormat(short index) =>
        _formatsByIndex.TryGetValue(index, out var fmt) ? fmt : null;

    internal void AddBiffFormat(short index, string format)
    {
        if (!_formatsByIndex.ContainsKey(index))
        {
            _formatsByIndex[index] = format;
            _formatsByString[format] = index;
            if (index >= _nextCustomIndex) _nextCustomIndex = (short)(index + 1);
        }
    }

    internal IEnumerable<KeyValuePair<short, string>> GetUserDefinedFormats() =>
        _formatsByIndex.Where(kv => kv.Key >= 164);
}
