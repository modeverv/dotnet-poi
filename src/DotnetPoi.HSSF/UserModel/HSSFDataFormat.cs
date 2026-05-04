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
        ArgumentNullException.ThrowIfNull(format);
        if (_formatsByString.TryGetValue(format, out var existing))
        {
            return existing;
        }

        var index = _nextCustomIndex++;
        _formatsByString[format] = index;
        _formatsByIndex[index] = format;
        return index;
    }

    public string? getFormat(short index) => _formatsByIndex.GetValueOrDefault(index);
}
