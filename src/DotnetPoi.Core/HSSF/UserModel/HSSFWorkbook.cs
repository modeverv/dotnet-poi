using DotnetPoi.HSSF.Record;
using DotnetPoi.POIFS.Crypt;
using DotnetPoi.SS.UserModel;

namespace DotnetPoi.HSSF.UserModel;

public sealed class HSSFWorkbook : IWorkbook
{
    private readonly List<HSSFSheet> _sheets = new();
    private readonly List<HSSFFont> _fonts = new();
    private readonly List<HSSFCellStyle> _cellStyles = new();
    private HSSFCreationHelper? _creationHelper;
    private HSSFDataFormat? _dataFormat;

    public HSSFWorkbook()
    {
        _fonts.Add(new HSSFFont(0));
        _cellStyles.Add(new HSSFCellStyle(this, 0));
    }

    public HSSFWorkbook(Stream stream)
        : this()
    {
        ArgumentNullException.ThrowIfNull(stream);
        Load(stream);
    }

    public HSSFSheet createSheet() => createSheet("Sheet" + (_sheets.Count + 1));

    public HSSFSheet createSheet(string sheetname)
    {
        if (string.IsNullOrEmpty(sheetname))
        {
            throw new ArgumentException("Sheet name must not be empty.", nameof(sheetname));
        }

        var sheet = new HSSFSheet(this, sheetname);
        _sheets.Add(sheet);
        return sheet;
    }

    public HSSFSheet getSheetAt(int index) => _sheets[index];

    public HSSFSheet? getSheet(string name) =>
        _sheets.FirstOrDefault(sheet => string.Equals(sheet.SheetName, name, StringComparison.Ordinal));

    public int getNumberOfSheets() => _sheets.Count;

    public HSSFCreationHelper getCreationHelper() => _creationHelper ??= new HSSFCreationHelper(this);

    public HSSFCellStyle createCellStyle()
    {
        var style = new HSSFCellStyle(this, _cellStyles.Count);
        _cellStyles.Add(style);
        return style;
    }

    public HSSFCellStyle getCellStyleAt(int idx) => _cellStyles[idx];

    public HSSFDataFormat createDataFormat() => _dataFormat ??= new HSSFDataFormat();

    public HSSFFont createFont()
    {
        var font = new HSSFFont(_fonts.Count);
        _fonts.Add(font);
        return font;
    }

    public HSSFFont getFontAt(int idx) => _fonts[idx];

    public int addPicture(byte[] pictureData, int format)
    {
        // TODO: [dotnet-poi] Not yet ported
        // Original: poi/poi/src/main/java/org/apache/poi/hssf/usermodel/HSSFWorkbook.java#addPicture
        // Reason: Phase 6 xls bootstrap only persists basic cell values.
        // Issue: Phase 6 HSSF pictures backlog
        throw new NotImplementedException("HSSF pictures are not yet ported. See Phase 6 HSSF pictures backlog.");
    }

    public int addPicture(Stream stream, int format)
    {
        ArgumentNullException.ThrowIfNull(stream);
        using var memory = new MemoryStream();
        stream.CopyTo(memory);
        return addPicture(memory.ToArray(), format);
    }

    public IReadOnlyList<IPictureData> getAllPictures() => Array.Empty<IPictureData>();

    public void setForceFormulaRecalculation(bool value)
    {
        // TODO: [dotnet-poi] Not yet ported
        // Original: poi/poi/src/main/java/org/apache/poi/hssf/usermodel/HSSFWorkbook.java#setForceFormulaRecalculation
        // Reason: HSSF calculation settings are outside this Phase 6 xls bootstrap slice.
        // Issue: Phase 6 HSSF calc settings backlog
    }

    public bool getForceFormulaRecalculation() => false;

    public void write(Stream stream)
    {
        ArgumentNullException.ThrowIfNull(stream);
        var workbookStream = Biff8Workbook.WriteWorkbook(_sheets);
        CompoundFile.Write(stream, new Dictionary<string, byte[]>(StringComparer.Ordinal)
        {
            ["Workbook"] = workbookStream
        });
    }

    public void close()
    {
    }

    public void Dispose() => close();

    private void Load(Stream stream)
    {
        var streams = CompoundFile.ReadStreams(stream);
        if (!streams.TryGetValue("Workbook", out var workbookStream) &&
            !streams.TryGetValue("Book", out workbookStream) &&
            !streams.TryGetValue("WORKBOOK", out workbookStream))
        {
            throw new InvalidDataException("The OLE2 document does not contain a Workbook stream.");
        }

        _sheets.Clear();
        Biff8Workbook.ReadWorkbook(workbookStream, this);
    }

    ISheet IWorkbook.createSheet() => createSheet();

    ISheet IWorkbook.createSheet(string sheetname) => createSheet(sheetname);

    ISheet IWorkbook.getSheetAt(int index) => getSheetAt(index);

    ISheet? IWorkbook.getSheet(string name) => getSheet(name);

    ICreationHelper IWorkbook.getCreationHelper() => getCreationHelper();

    ICellStyle IWorkbook.createCellStyle() => createCellStyle();

    ICellStyle IWorkbook.getCellStyleAt(int idx) => getCellStyleAt(idx);

    IDataFormat IWorkbook.createDataFormat() => createDataFormat();

    IFont IWorkbook.createFont() => createFont();

    IFont IWorkbook.getFontAt(int idx) => getFontAt(idx);

    internal IReadOnlyList<HSSFSheet> Sheets => _sheets;
}
