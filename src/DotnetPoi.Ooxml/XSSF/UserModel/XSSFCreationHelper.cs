using DotnetPoi.SS.UserModel;

namespace DotnetPoi.XSSF.UserModel;

public sealed class XSSFCreationHelper : ICreationHelper
{
    private readonly XSSFWorkbook _workbook;

    // Factory delegate for optional formula evaluator registration.
    // Set via RegisterFormulaEvaluatorFactory (typically from DotnetPoi.Formula).
    private static Func<IWorkbook, IFormulaEvaluator>? _formulaEvaluatorFactory;

    /// <summary>
    /// Registers a factory that creates IFormulaEvaluator instances.
    /// Called automatically by DotnetPoi.Formula when it is referenced.
    /// </summary>
    public static void RegisterFormulaEvaluatorFactory(Func<IWorkbook, IFormulaEvaluator> factory)
    {
        _formulaEvaluatorFactory = factory;
    }

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

    public XSSFRichTextString createRichTextString(string text)
    {
        return new XSSFRichTextString(text);
    }

    public IFormulaEvaluator createFormulaEvaluator()
    {
        if (_formulaEvaluatorFactory != null)
            return _formulaEvaluatorFactory(_workbook);

        // Lazy discovery: if DotnetPoi.Formula is referenced at runtime,
        // loading the assembly triggers its module initializer / static constructor.
        TryAutoRegisterFactory();

        if (_formulaEvaluatorFactory != null)
            return _formulaEvaluatorFactory(_workbook);

        throw new NotSupportedException(
            "Formula evaluation requires the DotnetPoi.Formula NuGet package. " +
            "Install it and construct 'new FormulaEvaluator(workbook)' directly.");
    }

    private static void TryAutoRegisterFactory()
    {
        try
        {
            var evaluatorType = Type.GetType(
                "DotnetPoi.Formula.FormulaEvaluator, DotnetPoi.Formula",
                throwOnError: false);
            if (evaluatorType != null)
            {
                System.Runtime.CompilerServices.RuntimeHelpers.RunClassConstructor(
                    evaluatorType.TypeHandle);
            }
        }
        catch
        {
            // DotnetPoi.Formula is not available; factory stays null.
        }
    }

    IDataFormat ICreationHelper.createDataFormat() => createDataFormat();

    IFormulaEvaluator ICreationHelper.createFormulaEvaluator() => createFormulaEvaluator();
}
