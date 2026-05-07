using DotnetPoi.SS.UserModel;

namespace DotnetPoi.HSSF.UserModel;

public sealed class HSSFCreationHelper : ICreationHelper
{
    private readonly HSSFWorkbook _workbook;

    internal HSSFCreationHelper(HSSFWorkbook workbook)
    {
        _workbook = workbook;
    }

    public HSSFDataFormat createDataFormat() => _workbook.createDataFormat();

    IDataFormat ICreationHelper.createDataFormat() => createDataFormat();

    public IFormulaEvaluator createFormulaEvaluator()
    {
        // TODO: [dotnet-poi] Not yet ported
        // Original: poi/poi/src/main/java/org/apache/poi/hssf/usermodel/HSSFFormulaEvaluator.java
        // Reason: Formula evaluation remains outside this Phase 6 xls bootstrap slice.
        // Issue: Phase 6 HSSF formula evaluator backlog
        throw new NotImplementedException("HSSF formula evaluation is not yet ported. See Phase 6 HSSF formula evaluator backlog.");
    }
}
