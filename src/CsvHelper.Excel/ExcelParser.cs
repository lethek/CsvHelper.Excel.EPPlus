using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Threading.Tasks;

using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;


namespace CsvHelper.Excel
{

    using System;
    using System.Linq;

    using Configuration;


    /// <summary>
    /// Parses an Excel file.
    /// </summary>
    public class ExcelParser : IParser
    {

        /// <summary>
        /// Creates a new parser using a new <see cref="ExcelPackage"/> from the given <paramref name="path"/> and uses the given <paramref name="configuration"/>.
        /// </summary>
        /// <param name="path">The path.</param>
        /// <param name="configuration">The configuration.</param>
        public ExcelParser(string path, CsvConfiguration configuration = null)
            : this(new ExcelPackage(new FileInfo(path)), configuration)
        {
            _shouldDisposeWorkbook = true;
        }


        /// <summary>
        /// Creates a new parser using a new <see cref="ExcelPackage"/> from the given <paramref name="path"/> and uses the given <paramref name="configuration"/>.
        /// </summary>
        /// <param name="path">The path to the workbook.</param>
        /// <param name="sheetName">The name of the sheet to import data from.</param>
        /// <param name="configuration">The configuration.</param>
        public ExcelParser(string path, string sheetName, CsvConfiguration configuration = null)
            : this(new ExcelPackage(new FileInfo(path)), sheetName, configuration)
        {
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


        private ExcelParser(ExcelRangeBase range, CsvConfiguration configuration)
        {
            Workbook = range.Worksheet.Workbook;
            this._range = range;
            Configuration = configuration ?? new CsvConfiguration(CultureInfo.CurrentCulture);
            Context = new ReadingContext(TextReader.Null, Configuration, false);
            FieldCount = range.Worksheet.Dimension.Columns;
        }


        /// <summary>
        /// Gets the reading context
        /// </summary>
        public ReadingContext Context { get; }

        /// <summary>
        /// Gets the configuration.
        /// </summary>
        public CsvConfiguration Configuration { get; }

        /// <summary>
        /// Gets the filed reader
        /// </summary>
        public IFieldReader FieldReader { get; }

        /// <summary>
        /// Gets the workbook from which we are reading data.
        /// </summary>
        /// <value>
        /// The workbook.
        /// </value>
        public ExcelWorkbook Workbook { get; }

        /// <summary>
        /// Gets the field count.
        /// </summary>
        public int FieldCount { get; }

        /// <summary>
        /// Gets the row of the Excel file that the parser is currently on.
        /// </summary>
        public int Row { get; private set; } = 1;

        /// <summary>
        /// Gets and sets the number of rows to offset the start position from.
        /// </summary>
        public int RowOffset { get; set; } = 0;

        /// <summary>
        /// Gets and sets the number of columns to offset the start position from.
        /// </summary>
        public int ColumnOffset { get; set; } = 0;


        /// <summary>
        /// Reads a record from the Excel file.
        /// </summary>
        /// <returns>
        /// A <see cref="T:String[]" /> of fields for the record read.
        /// </returns>
        /// <exception cref="ObjectDisposedException">Thrown if the parser has been disposed.</exception>
        public virtual string[] Read()
        {
            CheckDisposed();

            var fromRow = _range.Start.Row + Row + RowOffset - 1;
            var toRow = _range.Start.Row + Row + RowOffset - 1;
            var fromColumn = _range.Start.Column + ColumnOffset;
            var toColumn = _range.Start.Column + ColumnOffset + FieldCount - 1;

            var subRange = _range.Worksheet.Cells[fromRow, fromColumn, toRow, toColumn];
            subRange.Calculate(DefaultExcelCalculationOption);

            int expectIndex = 0;
            var values = new List<string>();
            foreach (var cell in subRange) {
                //If the current cell is further ahead than expected then OpenOfficeXml has skipped 1 or more empty cells: insert nulls for those
                int actualIndex = (cell.Start.Row - subRange.Start.Row) * FieldCount + (cell.Start.Column - subRange.Start.Column);
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
            if (values.Count < FieldCount) {
                values.AddRange(Enumerable.Repeat((string)null, FieldCount - values.Count));
            }

            Row++;
            return values.ToArray();
        }


        /// <summary>
        /// Pretends to asynchronously read a record from the Excel file.
        /// </summary>
        /// <returns>
        ///  A <see cref="T:String[]" /> of fields for the record read.
        /// </returns>
        /// <exception cref="ObjectDisposedException">Thrown if the parser has been disposed.</exception>
        public Task<string[]> ReadAsync()
            => Task.FromResult(Read());


        IParserConfiguration IParser.Configuration => Configuration;


        /// <summary>
        /// Performs application-defined tasks associated with freeing, releasing, or resetting unmanaged resources.
        /// </summary>
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }


        /// <summary>
        /// Finalizes an instance of the <see cref="ExcelParser"/> class.
        /// </summary>
        ~ExcelParser()
        {
            Dispose(false);
        }


        /// <summary>
        /// Releases unmanaged and - optionally - managed resources.
        /// </summary>
        /// <param name="disposing"><c>true</c> to release both managed and unmanaged resources; <c>false</c> to release only unmanaged resources.</param>
        protected virtual void Dispose(bool disposing)
        {
            if (_isDisposed) return;
            if (disposing) {
                if (_shouldDisposeWorkbook) Workbook.Dispose();
            }

            _isDisposed = true;
        }


        /// <summary>
        /// Checks if the instance has been disposed of.
        /// </summary>
        /// <exception cref="ObjectDisposedException" />
        protected virtual void CheckDisposed()
        {
            if (_isDisposed) {
                throw new ObjectDisposedException(GetType().ToString());
            }
        }


        private readonly bool _shouldDisposeWorkbook;
        private readonly ExcelRangeBase _range;
        private bool _isDisposed;

        private static readonly ExcelCalculationOption DefaultExcelCalculationOption = new ExcelCalculationOption();
    }

}
