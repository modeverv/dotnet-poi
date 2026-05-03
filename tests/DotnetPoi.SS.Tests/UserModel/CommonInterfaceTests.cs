using DotnetPoi.SS.UserModel;
using DotnetPoi.XSSF.UserModel;
using Xunit;

namespace DotnetPoi.SS.Tests.UserModel;

public class CommonInterfaceTests
{
    [Fact]
    public void XSSFWorkbook_ImplementsIWorkbook()
    {
        using IWorkbook workbook = new XSSFWorkbook();
        Assert.IsAssignableFrom<IWorkbook>(workbook);
    }

    [Fact]
    public void IWorkbook_CreateSheet_ReturnsISheet()
    {
        using IWorkbook workbook = new XSSFWorkbook();

        ISheet sheet = workbook.createSheet("Test");

        Assert.NotNull(sheet);
        Assert.IsAssignableFrom<ISheet>(sheet);
    }

    [Fact]
    public void ISheet_CreateRow_ReturnsIRow()
    {
        using IWorkbook workbook = new XSSFWorkbook();
        ISheet sheet = workbook.createSheet("Test");

        IRow row = sheet.createRow(0);

        Assert.NotNull(row);
        Assert.Equal(0, row.getRowNum());
        Assert.Same(sheet, row.getSheet());
    }

    [Fact]
    public void IRow_CreateCell_ReturnsICell()
    {
        using IWorkbook workbook = new XSSFWorkbook();
        IRow row = workbook.createSheet("Test").createRow(0);

        ICell cell = row.createCell(0);

        Assert.NotNull(cell);
        Assert.Equal(0, cell.getColumnIndex());
        Assert.Equal(CellType.Blank, cell.getCellType());
    }

    [Fact]
    public void ICell_SetStringValue_ReturnsCorrectType()
    {
        using IWorkbook workbook = new XSSFWorkbook();
        ICell cell = workbook.createSheet("Test").createRow(0).createCell(0);

        cell.setCellValue("hello");

        Assert.Equal(CellType.String, cell.getCellType());
        Assert.Equal("hello", cell.getStringCellValue());
    }

    [Fact]
    public void ICell_SetNumericValue_ReturnsCorrectType()
    {
        using IWorkbook workbook = new XSSFWorkbook();
        ICell cell = workbook.createSheet("Test").createRow(0).createCell(0);

        cell.setCellValue(42.5);

        Assert.Equal(CellType.Numeric, cell.getCellType());
        Assert.Equal(42.5, cell.getNumericCellValue());
    }

    [Fact]
    public void IWorkbook_CreateFont_ReturnsIFont()
    {
        using IWorkbook workbook = new XSSFWorkbook();

        IFont font = workbook.createFont();
        font.setBold(true);
        font.setFontName("Arial");

        Assert.True(font.getBold());
        Assert.Equal("Arial", font.getFontName());
    }

    [Fact]
    public void IWorkbook_CreateCellStyle_ReturnsICellStyle()
    {
        using IWorkbook workbook = new XSSFWorkbook();
        IFont font = workbook.createFont();

        ICellStyle style = workbook.createCellStyle();
        style.setFont(font);
        style.setFillPattern(FillPatternType.SolidForeground);
        style.setFillForegroundColor((short)IndexedColors.Yellow);
        style.setBorderBottom(BorderStyle.Thin);

        Assert.Same(font, style.getFont());
        Assert.Equal(FillPatternType.SolidForeground, style.getFillPattern());
        Assert.Equal(BorderStyle.Thin, style.getBorderBottom());
    }

    [Fact]
    public void ICell_SetCellStyle_ViaInterface()
    {
        using IWorkbook workbook = new XSSFWorkbook();
        ICellStyle style = workbook.createCellStyle();
        style.setDataFormat(workbook.createDataFormat().getFormat("0.00"));
        ICell cell = workbook.createSheet("Test").createRow(0).createCell(0);

        cell.setCellStyle(style);

        Assert.Same(style, cell.getCellStyle());
    }

    [Fact]
    public void IWorkbook_Write_ProducesReadableWorkbook()
    {
        using var ms = new MemoryStream();

        using (IWorkbook workbook = new XSSFWorkbook())
        {
            ISheet sheet = workbook.createSheet("Iface");
            IRow row = sheet.createRow(0);
            row.createCell(0).setCellValue("interface test");
            row.createCell(1).setCellValue(99.0);
            workbook.write(ms);
        }

        ms.Position = 0;
        using IWorkbook loaded = new XSSFWorkbook(ms);
        ISheet loadedSheet = loaded.getSheet("Iface")!;
        Assert.NotNull(loadedSheet);
        Assert.Equal("interface test", loadedSheet.getRow(0)!.getCell(0)!.getStringCellValue());
        Assert.Equal(99.0, loadedSheet.getRow(0)!.getCell(1)!.getNumericCellValue());
    }

    [Fact]
    public void IWorkbook_GetSheet_ReturnsNullForMissingSheet()
    {
        using IWorkbook workbook = new XSSFWorkbook();
        workbook.createSheet("Exists");

        Assert.Null(workbook.getSheet("DoesNotExist"));
        Assert.NotNull(workbook.getSheet("Exists"));
    }

    [Fact]
    public void IRow_GetCell_ReturnsNullForMissingCell()
    {
        using IWorkbook workbook = new XSSFWorkbook();
        IRow row = workbook.createSheet("Test").createRow(0);
        row.createCell(0).setCellValue("present");

        Assert.NotNull(row.getCell(0));
        Assert.Null(row.getCell(5));
    }

    [Fact]
    public void IWorkbook_GetCreationHelper_ReturnsICreationHelper()
    {
        using IWorkbook workbook = new XSSFWorkbook();

        ICreationHelper helper = workbook.getCreationHelper();

        Assert.NotNull(helper);
        Assert.NotNull(helper.createDataFormat());
    }
}
