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

    // Freeze panes
    void createFreezePane(int colSplit, int rowSplit);
    void createFreezePane(int colSplit, int rowSplit, int leftmostColumn, int topRow);

    // Hidden columns
    void setColumnHidden(int columnIndex, bool hidden);
    bool isColumnHidden(int columnIndex);
}
