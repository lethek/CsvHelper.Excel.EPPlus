using System.Globalization;
using System.Runtime.CompilerServices;

using CsvHelper.Configuration;

using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;


namespace CsvHelper.Excel.EPPlus;

/// <summary>
/// Parses an Excel file.
/// </summary>
public class ExcelParser : IParser
#if NETSTANDARD2_1_OR_GREATER
        , IAsyncDisposable
#endif
{
    /// <summary>
    /// Gets the workbook from which we are reading data.
    /// </summary>
    /// <value>
    /// The workbook.
    /// </value>
    public ExcelWorkbook Workbook { get; }


    /// <summary>
    /// Creates a new parser using a new <see cref="ExcelPackage"/> from the given <paramref name="path"/> and uses the given <paramref name="configuration"/>.
    /// </summary>
    /// <param name="path">The path to the workbook.</param>
    /// <param name="sheetName">The name of the sheet to import from. If null then the first worksheet in the workbook is used.</param>
    /// <param name="configuration">The configuration.</param>
    public ExcelParser(string path, string sheetName = null, IParserConfiguration configuration = null)
        : this(new ExcelPackage(new FileInfo(path)), sheetName, configuration, false) {
        _isPackageOwner = true;
    }


    /// <summary>
    /// Creates a new parser using a new <see cref="ExcelPackage"/> from the given <paramref name="stream"/> and uses the given <paramref name="configuration"/>.
    /// </summary>
    /// <param name="stream">The stream of the package.</param>
    /// <param name="sheetName">The name of the sheet to import from. If null then the first worksheet in the workbook is used.</param>
    /// <param name="configuration">The configuration.</param>
    /// <param name="leaveOpen">If set to <see langword="true"/>, disposing <see cref="ExcelParser"/> does not also automatically dispose the <paramref name="stream"/>.</param>
    public ExcelParser(Stream stream, string sheetName = null, IParserConfiguration configuration = null, bool leaveOpen = false)
        : this(new ExcelPackage(stream), sheetName, configuration, leaveOpen) {
        _isPackageOwner = true;
        _stream = stream;
    }


    /// <summary>
    /// Creates a new parser using the given <see cref="ExcelPackage"/> and <see cref="Configuration"/>.
    /// </summary>
    /// <param name="package">The <see cref="ExcelPackage"/> with the data.</param>
    /// <param name="sheetName">The name of the sheet to import from. If null then the first worksheet in the workbook is used.</param>
    /// <param name="configuration">The configuration.</param>
    /// <param name="leaveOpen">If set to <see langword="true"/>, disposing <see cref="ExcelParser"/> does not also automatically dispose the <paramref name="package"/>.</param>
    public ExcelParser(ExcelPackage package, string sheetName = null, IParserConfiguration configuration = null,
        bool leaveOpen = false)
        : this(package.Workbook, sheetName, configuration, leaveOpen)
    {
        _package = package;
    }


    /// <summary>
    /// Creates a new parser using the given <see cref="ExcelWorkbook"/> and <see cref="Configuration"/>.
    /// </summary>
    /// <param name="workbook">The <see cref="ExcelWorkbook"/> with the data.</param>
    /// <param name="sheetName">The name of the sheet to import from. If null then the first worksheet in the workbook is used.</param>
    /// <param name="configuration">The configuration.</param>
    /// <param name="leaveOpen">If set to <see langword="true"/>, disposing <see cref="ExcelParser"/> does not also automatically dispose the <paramref name="workbook"/>.</param>
    public ExcelParser(ExcelWorkbook workbook, string sheetName = null, IParserConfiguration configuration = null, bool leaveOpen = false)
        : this(
            sheetName != null ? workbook.Worksheets[sheetName] : workbook.Worksheets.First(),
            configuration,
            leaveOpen
        ) { }


    /// <summary>
    /// Creates a new parser using the given <see cref="ExcelWorksheet"/> and <see cref="Configuration"/>.
    /// </summary>
    /// <param name="worksheet">The <see cref="ExcelWorksheet"/> with the data.</param>
    /// <param name="configuration">The configuration.</param>
    /// <param name="leaveOpen">If set to <see langword="true"/>, disposing <see cref="ExcelParser"/> does not also automatically dispose the <see cref="ExcelWorkbook"/> which owns <paramref name="worksheet"/>.</param>
    public ExcelParser(ExcelWorksheet worksheet, IParserConfiguration configuration = null, bool leaveOpen = false)
        : this((ExcelRangeBase)worksheet.Cells, configuration, leaveOpen) { }


    /// <summary>
    /// Creates a new parser using the given <see cref="ExcelRange"/> and <see cref="Configuration"/>.
    /// </summary>
    /// <param name="range">The <see cref="ExcelRange"/> with the data.</param>
    /// <param name="configuration">The configuration.</param>
    /// <param name="leaveOpen">If set to <see langword="true"/>, disposing <see cref="ExcelParser"/> does not also automatically dispose the <see cref="ExcelWorkbook"/> which owns <paramref name="range"/>.</param>
    public ExcelParser(ExcelRange range, IParserConfiguration configuration = null, bool leaveOpen = false)
        : this((ExcelRangeBase)range, configuration, leaveOpen) { }


    private ExcelParser(ExcelRangeBase range, IParserConfiguration configuration, bool leaveOpen) {
        Configuration = configuration ?? new CsvConfiguration(CultureInfo.InvariantCulture);
        Context = new CsvContext(this);
        Workbook = range.Worksheet.Workbook;

        _range = (range.Address == "A:XFD")
            ? range.Worksheet.Cells[range.Worksheet.Dimension.Address]
            : range;

        _leaveOpen = leaveOpen;
    }


    /// <summary>
    /// Reads a record from the Excel file.
    /// </summary>
    /// <returns>
    /// A <see cref="T:String[]" /> of fields for the record read.
    /// </returns>
    /// <exception cref="ObjectDisposedException">Thrown if the parser has been disposed.</exception>
    public bool Read() {
        if (Row > RowCount) {
            return false;
        }

        Record = GetRecord();
        _row++;
        _rawRow++;
        return true;
    }


    /// <summary>
    /// Pretends to asynchronously read a record from the Excel file.
    /// </summary>
    /// <returns>
    ///  A <see cref="T:String[]" /> of fields for the record read.
    /// </returns>
    /// <exception cref="ObjectDisposedException">Thrown if the parser has been disposed.</exception>
    public Task<bool> ReadAsync()
        => Task.FromResult(Read());


    private string[] GetRecord() {
        var fromRow = _range.Start.Row + RowOffset + Row - 1;
        var fromColumn = _range.Start.Column + ColumnOffset;

        var toRow = fromRow;
        var toColumn = fromColumn + ColumnCount;

        var subRange = _range.Worksheet.Cells[fromRow, fromColumn, toRow, toColumn];
        subRange.Calculate(DefaultExcelCalculationOption);

        int expectedIndex = 0;
        var values = new List<string>(Count);
        foreach (var cell in subRange) {
            int actualIndex = (cell.Start.Row - subRange.Start.Row) * Count + (cell.Start.Column - subRange.Start.Column);

            //If the current cell is further ahead than expected then OpenOfficeXml has skipped 1 or more empty cells: insert nulls for those
            AddEmptyValuesForSkippedCells(values, actualIndex - expectedIndex);

            //Now we can add the value of the current cell
            values.Add(ApplyTrimOptions(cell.GetValue<string>()));

            expectedIndex = actualIndex + 1;
        }

        if (!values.Any()) {
            return null;
        }

        //If the number of values is fewer than expected then OpenOfficeXml has skipped 1 or more empty trailing cells: append nulls for those
        AddEmptyValuesForSkippedCells(values, Count - values.Count);

        return values.ToArray();
    }


    private string ApplyTrimOptions(string value)
        => Configuration.TrimOptions.HasFlag(TrimOptions.Trim) ? value?.Trim() : value;


    public long ByteCount => -1;
    public long CharCount => -1;
    public int Count => ColumnCount;

    public string this[int index] => Record.ElementAtOrDefault(index);

    public string[] Record { get; private set; }

    public string RawRecord => String.Join(Delimiter, Record);

    /// <summary>
    /// Gets the row of the Excel file that the parser is currently on.
    /// </summary>
    public int Row => _row;

    public int RawRow => _rawRow + _range.Start.Row - 1;

    /// <summary>
    /// Gets and sets the number of rows to offset the start position from.
    /// </summary>
    public int RowOffset { get; set; } = 0;

    /// <summary>
    /// Gets and sets the number of columns to offset the start position from.
    /// </summary>
    public int ColumnOffset { get; set; } = 0;

    public int RowCount => _range.Rows - RowOffset;

    public int ColumnCount => _range.Columns - ColumnOffset;
    
    /// <summary>
    /// Gets the reading context
    /// </summary>
    public CsvContext Context { get; }

    /// <summary>
    /// Gets the configuration.
    /// </summary>
    public IParserConfiguration Configuration { get; }

    public string Delimiter => Configuration.Delimiter;

    /// <summary>
    /// Performs application-defined tasks associated with freeing, releasing, or resetting unmanaged resources.
    /// </summary>
    public void Dispose() {
        Dispose(true);
        GC.SuppressFinalize(this);
    }


    /// <summary>
    /// Releases unmanaged and - optionally - managed resources.
    /// </summary>
    /// <param name="disposing"><c>true</c> to release both managed and unmanaged resources; <c>false</c> to release only unmanaged resources.</param>
    protected virtual void Dispose(bool disposing)
    {
        if (_disposed) {
            return;
        }
        if (disposing) {
            if (!_leaveOpen || _isPackageOwner) {
                Workbook.Dispose();
                _package?.Dispose();
            }
            if (!_leaveOpen) {
                _stream?.Dispose();
            }
        }
        _disposed = true;
    }


#if NETSTANDARD2_1_OR_GREATER
        /// <inheritdoc/>
        public async ValueTask DisposeAsync() {
            if (_disposed) {
                return;
            }
            if (!_leaveOpen || _isPackageOwner) {
                Workbook.Dispose();
                _package?.Dispose();
            }
            if (!_leaveOpen) {
                if (_stream != null) {
                    await _stream.DisposeAsync().ConfigureAwait(false);
                }
            }
            _disposed = true;
        }
#endif


    private static void AddEmptyValuesForSkippedCells(List<string> list, int count) {
        if (count > 0) {
            list.AddRange(Enumerable.Repeat((string)null, count));
        }
    }


    private readonly bool _isPackageOwner;
    private readonly bool _leaveOpen;

    private readonly ExcelPackage _package;
    private readonly ExcelRangeBase _range;
    private readonly Stream _stream;

    private bool _disposed;

    private int _row = 1;
    private int _rawRow = 1;

    private static readonly ExcelCalculationOption DefaultExcelCalculationOption = new();
}