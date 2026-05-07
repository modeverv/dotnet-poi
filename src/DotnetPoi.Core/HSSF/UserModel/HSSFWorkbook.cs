using DotnetPoi.HSSF.Record;
using DotnetPoi.POIFS.Crypt;
using DotnetPoi.SS.UserModel;

namespace DotnetPoi.HSSF.UserModel;

public sealed class HSSFWorkbook : IWorkbook
{
    private readonly List<HSSFSheet> _sheets = new();
    private readonly List<HSSFFont> _fonts = new();
    private readonly List<HSSFCellStyle> _cellStyles = new();
    private CompoundFileDocument? _preservedOleDocument;
    private byte[]? _preservedWorkbookStream;
    private string _workbookStreamName = "Workbook";
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
        Guard.ThrowIfNull(stream, nameof(stream));
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
        Guard.ThrowIfNull(stream, nameof(stream));
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

    public void protectWorkbook(bool protect) { }
    public bool isWorkbookProtected() => false;

    // Active sheet / selected tab
    public void setActiveSheet(int index) { }
    public int getActiveSheetIndex() => 0;
    public void setSelectedTab(int index) { }

    public void write(Stream stream)
    {
        Guard.ThrowIfNull(stream, nameof(stream));
        var workbookStream = Biff8Workbook.WriteWorkbook(_sheets, _preservedWorkbookStream);
        var document = _preservedOleDocument is null
            ? new CompoundFileDocument(new Dictionary<string, byte[]>(StringComparer.Ordinal))
            : new CompoundFileDocument(_preservedOleDocument.Streams, _preservedOleDocument.EntryMetadata);

        RemoveWorkbookStreamAliases(document.Streams);
        document.Streams[_workbookStreamName] = workbookStream;
        CompoundFile.Write(stream, document);
    }

    public void close()
    {
    }

    public void Dispose() => close();

    private void Load(Stream stream)
    {
        var document = CompoundFile.ReadDocument(stream);
        if (!TryGetWorkbookStream(document.Streams, out var workbookStream, out var workbookStreamName))
        {
            throw new InvalidDataException("The OLE2 document does not contain a Workbook stream.");
        }

        _preservedOleDocument = document;
        _preservedWorkbookStream = workbookStream.ToArray();
        _workbookStreamName = workbookStreamName;
        _sheets.Clear();
        BeginBiffLoad();
        Biff8Workbook.ReadWorkbook(workbookStream, this);
        if (_cellStyles.Count == 0) _cellStyles.Add(new HSSFCellStyle(this, 0));
        if (_fonts.Count == 0) _fonts.Add(new HSSFFont(0));
    }

    private static bool TryGetWorkbookStream(
        IReadOnlyDictionary<string, byte[]> streams,
        out byte[] workbookStream,
        out string streamName)
    {
        foreach (var candidate in WorkbookStreamAliases)
        {
            if (streams.TryGetValue(candidate, out workbookStream!))
            {
                streamName = candidate;
                return true;
            }
        }

        workbookStream = Array.Empty<byte>();
        streamName = "Workbook";
        return false;
    }

    private static void RemoveWorkbookStreamAliases(IDictionary<string, byte[]> streams)
    {
        foreach (var candidate in WorkbookStreamAliases)
        {
            streams.Remove(candidate);
        }
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

    internal void BeginBiffLoad()
    {
        _fonts.Clear();
        _cellStyles.Clear();
    }

    internal void AddFontFromBiff(HSSFFont font) => _fonts.Add(font);

    internal void AddStyleFromBiff(HSSFCellStyle style) => _cellStyles.Add(style);

    internal int getNumberOfFonts() => _fonts.Count;

    internal int getNumberOfCellStyles() => _cellStyles.Count;

    internal HSSFDataFormat GetOrCreateDataFormat() => createDataFormat();

    private static readonly string[] WorkbookStreamAliases =
    {
        "Workbook",
        "Book",
        "WORKBOOK",
        "BOOK"
    };
}
