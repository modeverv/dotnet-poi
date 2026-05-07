namespace DotnetPoi.SS.UserModel;

public interface IWorkbook : IDisposable
{
    ISheet createSheet();
    ISheet createSheet(string sheetname);
    ISheet getSheetAt(int index);
    ISheet? getSheet(string name);
    int getNumberOfSheets();
    ICreationHelper getCreationHelper();
    ICellStyle createCellStyle();
    ICellStyle getCellStyleAt(int idx);
    IDataFormat createDataFormat();
    IFont createFont();
    IFont getFontAt(int idx);
    int addPicture(byte[] pictureData, int format);
    int addPicture(Stream stream, int format);
    IReadOnlyList<IPictureData> getAllPictures();
    void setForceFormulaRecalculation(bool value);
    bool getForceFormulaRecalculation();
    void write(Stream stream);
    void close();

    // Workbook protection
    void protectWorkbook(bool protect);
    bool isWorkbookProtected();

    // Active sheet / selected tab
    void setActiveSheet(int index);
    int getActiveSheetIndex();
    void setSelectedTab(int index);
}
