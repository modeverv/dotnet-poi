namespace DotnetPoi.SS.UserModel;

public interface IRow
{
    ICell createCell(int columnIndex);
    ICell? getCell(int cellnum);
    short getLastCellNum();
    int getRowNum();
    ISheet getSheet();
}
