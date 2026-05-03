namespace DotnetPoi.SS.UserModel;

public interface ICell
{
    int getColumnIndex();
    int getRowIndex();
    CellType getCellType();
    void setCellValue(string? value);
    void setCellValue(double value);
    string getStringCellValue();
    double getNumericCellValue();
    ICellStyle getCellStyle();
    void setCellStyle(ICellStyle? style);
    ISheet getSheet();
    IRow getRow();
}
