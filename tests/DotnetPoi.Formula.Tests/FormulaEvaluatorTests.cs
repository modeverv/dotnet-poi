using DotnetPoi.Formula;
using DotnetPoi.SS.UserModel;
using DotnetPoi.XSSF.UserModel;
using Xunit;

namespace DotnetPoi.Formula.Tests;

public class FormulaEvaluatorTests
{
    [Fact]
    public void FormulaAssembly_DoesNotReferenceFormatImplementations()
    {
        var referencedAssemblies = typeof(FormulaEvaluator)
            .Assembly
            .GetReferencedAssemblies()
            .Select(a => a.Name)
            .ToHashSet(StringComparer.Ordinal);

        Assert.Contains("DotnetPoi.Common", referencedAssemblies);
        Assert.DoesNotContain("DotnetPoi.Core", referencedAssemblies);
        Assert.DoesNotContain("DotnetPoi.Ooxml", referencedAssemblies);
        Assert.DoesNotContain("DotnetPoi.Legacy", referencedAssemblies);
    }

    [Fact]
    public void EvaluateFormulaCell_Sum_StoresCachedNumericValue()
    {
        using var workbook = new XSSFWorkbook();
        var sheet = workbook.createSheet("Eval");
        var row = sheet.createRow(0);
        row.createCell(0).setCellValue(10.0);
        row.createCell(1).setCellValue(20.0);
        row.createCell(2).setCellValue(30.0);

        var sum = sheet.createRow(1).createCell(0);
        sum.setCellFormula("SUM(A1:C1)");

        var evaluator = new FormulaEvaluator(workbook);
        Assert.Equal(CellType.Numeric, evaluator.evaluateFormulaCell(sum));
        Assert.Equal(60.0, sum.getNumericCellValue());
    }

    [Fact]
    public void EvaluateFormulaCell_Average_StoresCachedNumericValue()
    {
        using var workbook = new XSSFWorkbook();
        var sheet = workbook.createSheet("Eval");
        var row = sheet.createRow(0);
        row.createCell(0).setCellValue(10.0);
        row.createCell(1).setCellValue(20.0);
        row.createCell(2).setCellValue(30.0);

        var avg = sheet.createRow(1).createCell(0);
        avg.setCellFormula("AVERAGE(A1:C1)");

        var evaluator = new FormulaEvaluator(workbook);
        Assert.Equal(CellType.Numeric, evaluator.evaluateFormulaCell(avg));
        Assert.Equal(20.0, avg.getNumericCellValue());
    }

    [Fact]
    public void EvaluateAll_RepresentativeFunctions_StoresCachedValues()
    {
        using var workbook = new XSSFWorkbook();
        var sheet = workbook.createSheet("EvalAll");
        var row = sheet.createRow(0);
        row.createCell(0).setCellValue(5.0);
        row.createCell(1).setCellValue(9.0);
        row.createCell(2).setCellValue("text");

        var min = sheet.createRow(1).createCell(0);
        min.setCellFormula("MIN(A1:B1)");
        var max = sheet.getRow(1)!.createCell(1);
        max.setCellFormula("MAX(A1:B1)");
        var count = sheet.getRow(1)!.createCell(2);
        count.setCellFormula("COUNT(A1:C1)");
        var text = sheet.getRow(1)!.createCell(3);
        text.setCellFormula("CONCATENATE(\"total=\",A1+B1)");

        new FormulaEvaluator(workbook).evaluateAll();

        Assert.Equal(5.0, min.getNumericCellValue());
        Assert.Equal(9.0, max.getNumericCellValue());
        Assert.Equal(2.0, count.getNumericCellValue());
        Assert.Equal("total=14", text.getStringCellValue());
    }

    [Fact]
    public void Evaluate_ArithmeticExpression_ReturnsCorrectValue()
    {
        using var workbook = new XSSFWorkbook();
        var sheet = workbook.createSheet("Arith");
        var row = sheet.createRow(0);
        row.createCell(0).setCellValue(10.0);
        row.createCell(1).setCellValue(20.0);

        var formula = row.createCell(2);
        formula.setCellFormula("A1+B1*2");

        var evaluator = new FormulaEvaluator(workbook);
        var result = evaluator.evaluate(formula);
        Assert.Equal(CellType.Numeric, result.getCellType());
        Assert.Equal(50.0, result.getNumberValue());
    }

    [Fact]
    public void EvaluateInCell_StringConcat_SetsCellValue()
    {
        using var workbook = new XSSFWorkbook();
        var sheet = workbook.createSheet("Concat");
        var row = sheet.createRow(0);
        row.createCell(0).setCellValue("hello");
        row.createCell(1).setCellValue("world");

        var concat = row.createCell(2);
        concat.setCellFormula("CONCATENATE(A1,\" \",B1)");

        var evaluator = new FormulaEvaluator(workbook);
        evaluator.evaluateInCell(concat);
        Assert.Equal(CellType.String, concat.getCellType());
        Assert.Equal("hello world", concat.getStringCellValue());
    }

    [Fact]
    public void Evaluate_ThroughInterface_CreatesEvaluator()
    {
        using var workbook = new XSSFWorkbook();
        var helper = workbook.getCreationHelper();
        var evaluator = helper.createFormulaEvaluator();
        Assert.NotNull(evaluator);
        Assert.IsAssignableFrom<IFormulaEvaluator>(evaluator);
    }

    [Fact]
    public void EvaluateFormulaCell_DivideByZero_ReturnsError()
    {
        using var workbook = new XSSFWorkbook();
        var sheet = workbook.createSheet("DivZero");
        var row = sheet.createRow(0);
        row.createCell(0).setCellValue(10.0);
        row.createCell(1).setCellValue(0.0);

        var div = row.createCell(2);
        div.setCellFormula("A1/B1");

        var evaluator = new FormulaEvaluator(workbook);
        var result = evaluator.evaluate(div);
        Assert.Equal(CellType.Error, result.getCellType());
    }

    [Fact]
    public void EvaluateFormulaCell_ReferenceToSelf_ReturnsError()
    {
        using var workbook = new XSSFWorkbook();
        var sheet = workbook.createSheet("Circular");
        var cell = sheet.createRow(0).createCell(0);
        cell.setCellFormula("A1");

        var evaluator = new FormulaEvaluator(workbook);
        var result = evaluator.evaluate(cell);
        Assert.Equal(CellType.Error, result.getCellType());
    }

    [Fact]
    public void Evaluate_RangeReference_ComputesSum()
    {
        using var workbook = new XSSFWorkbook();
        var sheet = workbook.createSheet("Range");
        var row = sheet.createRow(0);
        row.createCell(0).setCellValue(1.0);
        row.createCell(1).setCellValue(2.0);
        row.createCell(2).setCellValue(3.0);

        var sum = row.createCell(3);
        sum.setCellFormula("SUM(A1:C1)");

        var evaluator = new FormulaEvaluator(workbook);
        Assert.Equal(CellType.Numeric, evaluator.evaluateFormulaCell(sum));
        Assert.Equal(6.0, sum.getNumericCellValue());
    }

    [Fact]
    public void Evaluate_BooleanExpression_ReturnsBoolean()
    {
        using var workbook = new XSSFWorkbook();
        var sheet = workbook.createSheet("Bool");
        var cell = sheet.createRow(0).createCell(0);
        cell.setCellFormula("TRUE");

        var evaluator = new FormulaEvaluator(workbook);
        var result = evaluator.evaluate(cell);
        Assert.Equal(CellType.Boolean, result.getCellType());
        Assert.True(result.getBooleanValue());
    }
}
