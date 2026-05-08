using System.Reflection;
using DotnetPoi.HSSF.UserModel;
using DotnetPoi.SS.UserModel;
using DotnetPoi.XSLF.UserModel;
using DotnetPoi.XSSF.UserModel;
using DotnetPoi.XWPF.UserModel;
using Xunit;

namespace DotnetPoi.All.Tests;

public class ProjectShellTests
{
    [Theory]
    [InlineData("DotnetPoi.All")]
    [InlineData("DotnetPoi.Common")]
    [InlineData("DotnetPoi.POIFS")]
    [InlineData("DotnetPoi.Ooxml")]
    [InlineData("DotnetPoi.Legacy")]
    [InlineData("DotnetPoi.Formula")]
    public void PackageSurfaceAssembly_Loads(string assemblyName)
    {
        var assembly = Assembly.Load(assemblyName);

        Assert.Equal(assemblyName, assembly.GetName().Name);
    }

    [Fact]
    public void AllPackage_CanUseRepresentativeOoxmlLegacyAndFormulaSurfaces()
    {
        using var xssf = new XSSFWorkbook();
        var xssfSheet = xssf.createSheet("OOXML");
        var formulaCell = xssfSheet.createRow(0).createCell(0);
        formulaCell.setCellFormula("SUM(1,2)");
        xssf.getCreationHelper().createFormulaEvaluator().evaluateFormulaCell(formulaCell);
        Assert.Equal(3.0, formulaCell.getNumericCellValue());

        using var hssf = new HSSFWorkbook();
        hssf.createSheet("Legacy").createRow(0).createCell(0).setCellValue("xls");
        Assert.Equal("xls", hssf.getSheet("Legacy")!.getRow(0)!.getCell(0)!.getStringCellValue());

        using var xwpf = new XWPFDocument();
        xwpf.createParagraph().createRun().setText("docx");
        Assert.Contains("docx", xwpf.getParagraphs()[0].getText());

        using var xslf = new XMLSlideShow();
        Assert.Empty(xslf.getSlides());
    }
}
