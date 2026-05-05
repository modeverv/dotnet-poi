namespace DotnetPoi.SS.UserModel;

public interface ISheet
{
    IRow createRow(int rownum);
    IRow? getRow(int rownum);
    int getLastRowNum();
    IWorkbook getWorkbook();
}
