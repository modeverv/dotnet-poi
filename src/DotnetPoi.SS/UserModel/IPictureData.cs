namespace DotnetPoi.SS.UserModel;

public interface IPictureData
{
    int getPictureType();
    byte[] getData();
    string suggestFileExtension();
    string getMimeType();
}
