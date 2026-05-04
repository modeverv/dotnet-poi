namespace DotnetPoi.SS.UserModel;

/// <summary>
/// Evaluates formula cells.
/// Ported from org.apache.poi.ss.usermodel.FormulaEvaluator.
/// </summary>
public interface IFormulaEvaluator
{
    void clearAllCachedResultValues();
    void notifySetFormula(ICell cell);
    void notifyDeleteCell(ICell cell);
    void notifyUpdateCell(ICell cell);
    void evaluateAll();
    CellValue evaluate(ICell cell);
    CellType evaluateFormulaCell(ICell cell);
    ICell evaluateInCell(ICell cell);
}
