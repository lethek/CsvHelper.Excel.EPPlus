using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Threading.Tasks;

using CsvHelper.Configuration;

using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;


namespace CsvHelper.Excel
{

    /// <summary>
    /// Parses an Excel file.
    /// </summary>
    public class ExcelParser : IParser
    {
        /// <summary>
        /// Creates a new parser using a new <see cref="ExcelPackage"/> from the given <paramref name="stream"/> and uses the given <paramref name="configuration"/>.
        /// </summary>
        /// <param name="stream">The stream of the package.</param>
        /// <param name="configuration">The configuration.</param>
        public ExcelParser(Stream stream, CsvConfiguration configuration = null)
            : this(new ExcelPackage(stream), configuration) {
            _shouldDisposeWorkbook = true;
        }


        /// <summary>
        /// Creates a new parser using a new <see cref="ExcelPackage"/> from the given <paramref name="stream"/> and uses the given <paramref name="configuration"/>.
        /// </summary>
        /// <param name="stream">The stream of the package.</param>
        /// <param name="sheetName">The name of the sheet to import data from.</param>
        /// <param name="configuration">The configuration.</param>
        public ExcelParser(Stream stream, string sheetName, CsvConfiguration configuration = null)
            : this(new ExcelPackage(stream), sheetName, configuration) {
            _shouldDisposeWorkbook = true;
        }


        /// <summary>
        /// Creates a new parser using a new <see cref="ExcelPackage"/> from the given <paramref name="path"/> and uses the given <paramref name="configuration"/>.
        /// </summary>
        /// <param name="path">The path.</param>
        /// <param name="configuration">The configuration.</param>
        public ExcelParser(string path, CsvConfiguration configuration = null)
            : this(new ExcelPackage(new FileInfo(path)), configuration) {
            _shouldDisposeWorkbook = true;
        }


        /// <summary>
        /// Creates a new parser using a new <see cref="ExcelPackage"/> from the given <paramref name="path"/> and uses the given <paramref name="configuration"/>.
        /// </summary>
        /// <param name="path">The path to the workbook.</param>
        /// <param name="sheetName">The name of the sheet to import data from.</param>
        /// <param name="configuration">The configuration.</param>
        public ExcelParser(string path, string sheetName, CsvConfiguration configuration = null)
            : this(new ExcelPackage(new FileInfo(path)), sheetName, configuration) {
            _shouldDisposeWorkbook = true;
        }


        /// <summary>
        /// Creates a new parser using the given <see cref="ExcelPackage"/> and <see cref="Configuration"/>.
        /// <remarks>
        /// Will attempt to read the data from the first worksheet in the workbook.
        /// </remarks>
        /// </summary>
        /// <param name="package">The <see cref="ExcelPackage"/> with the data.</param>
        /// <param name="configuration">The configuration.</param>
        public ExcelParser(ExcelPackage package, CsvConfiguration configuration = null)
            : this(package.Workbook, configuration) { }


        /// <summary>
        /// Creates a new parser using the given <see cref="ExcelPackage"/> and <see cref="Configuration"/>.
        /// </summary>
        /// <param name="package">The <see cref="ExcelPackage"/> with the data.</param>
        /// <param name="sheetName">The name of the sheet to import from.</param>
        /// <param name="configuration">The configuration.</param>
        public ExcelParser(ExcelPackage package, string sheetName, CsvConfiguration configuration = null)
            : this(package.Workbook, sheetName, configuration) { }


        /// <summary>
        /// Creates a new parser using the given <see cref="ExcelWorkbook"/> and <see cref="Configuration"/>.
        /// <remarks>
        /// Will attempt to read the data from the first worksheet in the workbook.
        /// </remarks>
        /// </summary>
        /// <param name="workbook">The <see cref="ExcelWorkbook"/> with the data.</param>
        /// <param name="configuration">The configuration.</param>
        public ExcelParser(ExcelWorkbook workbook, CsvConfiguration configuration = null)
            : this(workbook.Worksheets.First(), configuration) { }


        /// <summary>
        /// Creates a new parser using the given <see cref="ExcelWorkbook"/> and <see cref="Configuration"/>.
        /// </summary>
        /// <param name="workbook">The <see cref="ExcelWorkbook"/> with the data.</param>
        /// <param name="sheetName">The name of the sheet to import from.</param>
        /// <param name="configuration">The configuration.</param>
        public ExcelParser(ExcelWorkbook workbook, string sheetName, CsvConfiguration configuration = null)
            : this(workbook.Worksheets[sheetName], configuration) { }


        /// <summary>
        /// Creates a new parser using the given <see cref="ExcelWorksheet"/> and <see cref="Configuration"/>.
        /// </summary>
        /// <param name="worksheet">The <see cref="ExcelWorksheet"/> with the data.</param>
        /// <param name="configuration">The configuration.</param>
        public ExcelParser(ExcelWorksheet worksheet, CsvConfiguration configuration = null)
            : this((ExcelRangeBase)worksheet.Cells, configuration) { }


        /// <summary>
        /// Creates a new parser using the given <see cref="ExcelRange"/> and <see cref="Configuration"/>.
        /// </summary>
        /// <param name="range">The <see cref="ExcelRange"/> with the data.</param>
        /// <param name="configuration">The configuration.</param>
        public ExcelParser(ExcelRange range, CsvConfiguration configuration = null)
            : this((ExcelRangeBase)range, configuration) { }


        private ExcelParser(ExcelRangeBase range, CsvConfiguration configuration) {
            Workbook = range.Worksheet.Workbook;
            _range = range;
            Configuration = configuration ?? new CsvConfiguration(CultureInfo.InvariantCulture);
            Context = new CsvContext(this);
            Count = range.Worksheet.Dimension.Columns;
        }


        /// <summary>
        /// Gets the workbook from which we are reading data.
        /// </summary>
        /// <value>
        /// The workbook.
        /// </value>
        public ExcelWorkbook Workbook { get; }


        /// <summary>
        /// Gets and sets the number of rows to offset the start position from.
        /// </summary>
        public int RowOffset { get; set; }

        /// <summary>
        /// Gets and sets the number of columns to offset the start position from.
        /// </summary>
        public int ColumnOffset { get; set; }


        /// <summary>
        /// Reads a record from the Excel file.
        /// </summary>
        /// <returns>
        /// A <see cref="T:String[]" /> of fields for the record read.
        /// </returns>
        /// <exception cref="ObjectDisposedException">Thrown if the parser has been disposed.</exception>
        public bool Read() {
            if (Row > _lastRow) {
                return false;
            }

            _currentRecord = GetRecord();
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


        IParserConfiguration IParser.Configuration => Configuration;

        public long ByteCount => -1;
        public long CharCount => -1;
        public int Count { get; }

        public string this[int index] => Record.ElementAtOrDefault(index);

        public string[] Record => _currentRecord;

        public string RawRecord => String.Join(Configuration.Delimiter, Record);

        /// <summary>
        /// Gets the row of the Excel file that the parser is currently on.
        /// </summary>
        public int Row => _row;

        public int RawRow => _rawRow;

        /// <summary>
        /// Gets the reading context
        /// </summary>
        public CsvContext Context { get; }

        /// <summary>
        /// Gets the configuration.
        /// </summary>
        public CsvConfiguration Configuration { get; }

        CsvContext IParser.Context => throw new NotImplementedException();


        /// <summary>
        /// Performs application-defined tasks associated with freeing, releasing, or resetting unmanaged resources.
        /// </summary>
        public void Dispose() {
            Dispose(true);
            GC.SuppressFinalize(this);
        }


        /// <summary>
        /// Finalizes an instance of the <see cref="ExcelParser"/> class.
        /// </summary>
        ~ExcelParser() {
            Dispose(false);
        }


        /// <summary>
        /// Releases unmanaged and - optionally - managed resources.
        /// </summary>
        /// <param name="disposing"><c>true</c> to release both managed and unmanaged resources; <c>false</c> to release only unmanaged resources.</param>
        protected virtual void Dispose(bool disposing) {
            if (!_disposed) {
                if (disposing) {
                    if (_shouldDisposeWorkbook) {
                        Workbook.Dispose();
                    }
                }
                _disposed = true;
            }
        }


        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        private string[] GetRecord() {
            /*var currentRow = _worksheet.Row(Row);
            var cells = currentRow.Cells(1, Count);
            var values = cells.Select(x => x.Value.ToString()).ToArray();
            return values;*/

            var fromRow = _range.Start.Row + Row + RowOffset - 1;
            var toRow = _range.Start.Row + Row + RowOffset - 1;
            var fromColumn = _range.Start.Column + ColumnOffset;
            var toColumn = _range.Start.Column + ColumnOffset + Count - 1;

            var subRange = _range.Worksheet.Cells[fromRow, fromColumn, toRow, toColumn];
            subRange.Calculate(DefaultExcelCalculationOption);

            int expectIndex = 0;
            var values = new List<string>(Count);
            foreach (var cell in subRange) {
                //If the current cell is further ahead than expected then OpenOfficeXml has skipped 1 or more empty cells: insert nulls for those
                int actualIndex = (cell.Start.Row - subRange.Start.Row) * Count + (cell.Start.Column - subRange.Start.Column);
                int indexDelta = actualIndex - expectIndex;
                if (indexDelta > 0) {
                    values.AddRange(Enumerable.Repeat((string)null, indexDelta));
                }

                //Now we can add the value of the current cell
                values.Add(cell.GetValue<string>());

                expectIndex = actualIndex + 1;
            }

            if (!values.Any()) {
                return null;
            }

            //If the number of values is fewer than expected then OpenOfficeXml has skipped 1 or more empty trailing cells: append nulls for those
            if (values.Count < Count) {
                values.AddRange(Enumerable.Repeat((string)null, Count - values.Count));
            }

            return values.ToArray();
        }


        private readonly bool _shouldDisposeWorkbook;
        private readonly ExcelRangeBase _range;
        private bool _disposed;

        private int _row = 1;
        private int _rawRow = 1;
        private int _lastRow;
        private string[] _currentRecord;

        //private readonly bool _leaveOpen;
        //private readonly Stream _stream;

        private static readonly ExcelCalculationOption DefaultExcelCalculationOption = new();
    }

}
