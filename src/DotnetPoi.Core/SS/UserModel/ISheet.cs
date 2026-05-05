using DotnetPoi.SS.Util;

namespace DotnetPoi.SS.UserModel;

public interface ISheet
{
    IRow createRow(int rownum);
    IRow? getRow(int rownum);
    int getLastRowNum();
    IWorkbook getWorkbook();

    // Merged cells
    void addMergedRegion(CellRangeAddress region);
    IReadOnlyList<CellRangeAddress> getMergedRegions();

    // Column width
    void setColumnWidth(int columnIndex, int width);
    int getColumnWidth(int columnIndex);
}
