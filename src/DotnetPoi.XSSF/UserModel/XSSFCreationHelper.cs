namespace DotnetPoi.XSSF.UserModel;

public sealed class XSSFCreationHelper
{
    private readonly XSSFWorkbook _workbook;

    internal XSSFCreationHelper(XSSFWorkbook workbook)
    {
        _workbook = workbook;
    }

    public XSSFWorkbook getWorkbook()
    {
        return _workbook;
    }
}
