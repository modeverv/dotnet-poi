namespace DotnetPoi.SS.UserModel;

public interface IDataFormat
{
    short getFormat(string format);
    string? getFormat(short index);
}
