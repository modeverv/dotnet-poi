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

    /// <summary>
    /// Returns the formula text for formula cells.
    /// Ported from org.apache.poi.ss.usermodel.Cell#getCellFormula.
    /// </summary>
    string? getCellFormula();

    void setCellValue(string? value);
    void setCellValue(double value);
    void setCellValue(bool value);
    void setCellFormula(string? formula);

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

    /// <summary>
    /// Sets the cached result of a formula evaluation on this cell.
    /// Only valid for formula cells (<see cref="CellType.Formula"/>).
    /// Ported from XSSFCell.SetFormulaCachedValue.
    /// </summary>
    void setCachedFormulaResult(CellValue value);

    /// <summary>
    /// Returns the hyperlink associated with this cell, or null if none.
    /// </summary>
    IHyperlink? getHyperlink();

    /// <summary>
    /// Assigns a hyperlink to this cell.
    /// </summary>
    void setHyperlink(IHyperlink? link);
}
