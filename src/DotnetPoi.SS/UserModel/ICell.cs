namespace DotnetPoi.SS.UserModel;

public interface ICell
{
    int getColumnIndex();
    int getRowIndex();
    CellType getCellType();

    /// <summary>
    /// Only valid for formula cells. Returns the cached result type:
    /// Numeric, String, Boolean, or Error.
    /// Ported from org.apache.poi.ss.usermodel.Cell#getCachedFormulaResultType.
    /// </summary>
    CellType getCachedFormulaResultType();

    void setCellValue(string? value);
    void setCellValue(double value);
    void setCellValue(bool value);

    string getStringCellValue();
    double getNumericCellValue();
    bool getBooleanCellValue();

    /// <summary>
    /// Returns the raw error string (e.g. "#DIV/0!") for error cells.
    /// Ported from org.apache.poi.xssf.usermodel.XSSFCell#getErrorCellString.
    /// </summary>
    string getErrorCellString();

    ICellStyle getCellStyle();
    void setCellStyle(ICellStyle? style);
    ISheet getSheet();
    IRow getRow();
}
