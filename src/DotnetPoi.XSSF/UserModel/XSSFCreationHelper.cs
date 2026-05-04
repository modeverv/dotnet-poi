using DotnetPoi.SS.UserModel;

namespace DotnetPoi.XSSF.UserModel;

public sealed class XSSFCreationHelper : ICreationHelper
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

    public XSSFDataFormat createDataFormat()
    {
        return _workbook.createDataFormat();
    }

    public XSSFClientAnchor createClientAnchor()
    {
        return new XSSFClientAnchor();
    }

    public XSSFFormulaEvaluator createFormulaEvaluator()
    {
        return new XSSFFormulaEvaluator(_workbook);
    }

    IDataFormat ICreationHelper.createDataFormat() => createDataFormat();

    IFormulaEvaluator ICreationHelper.createFormulaEvaluator() => createFormulaEvaluator();
}
