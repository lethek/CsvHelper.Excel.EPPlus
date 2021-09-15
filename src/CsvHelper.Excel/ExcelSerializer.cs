using System;
using System.Globalization;
using System.IO;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

using CsvHelper.Configuration;

using OfficeOpenXml;


namespace CsvHelper.Excel
{

    /// <summary>
    /// Defines methods used to serialize data into an Excel (2007+) file.
    /// </summary>
    public class ExcelSerializer : ISerializer
    {
        private readonly string path;
        private readonly bool disposePackage;
        private readonly ExcelRangeBase range;
        private bool disposed;
        private int currentRow = 1;


        /// <summary>
        /// Creates a new serializer using a new <see cref="ExcelPackage"/> saved to the given <paramref name="path"/>.
        /// <remarks>
        /// The package will not be saved until the serializer is disposed.
        /// </remarks>
        /// </summary>
        /// <param name="path">The path to which to save the package.</param>
        /// <param name="configuration">The configuration</param>
        public ExcelSerializer(string path, CsvConfiguration configuration = null)
            : this(new ExcelPackage(), configuration)
        {
            this.path = path;
            disposePackage = true;
        }


        /// <summary>
        /// Creates a new serializer using a new <see cref="ExcelPackage"/> saved to the given <paramref name="path"/>.
        /// <remarks>
        /// The package will not be saved until the serializer is disposed.
        /// </remarks>
        /// </summary>
        /// <param name="path">The path to which to save the package.</param>
        /// <param name="sheetName">The name of the sheet to which to save</param>
        public ExcelSerializer(string path, string sheetName)
            : this(new ExcelPackage(), sheetName)
        {
            this.path = path;
            disposePackage = true;
        }


        /// <summary>
        /// Creates a new serializer using the given <see cref="ExcelPackage"/> and <see cref="Configuration"/>.
        /// <remarks>
        /// The <paramref name="package"/> will <b><i>not</i></b> be disposed of when the serializer is disposed.
        /// The package will <b><i>not</i></b> be saved by the serializer.
        /// A new worksheet will be added to the package.
        /// </remarks>
        /// </summary>
        /// <param name="package">The package to write the data to.</param>
        /// <param name="configuration">The configuration.</param>
        public ExcelSerializer(ExcelPackage package, CsvConfiguration configuration = null)
            : this(package, "Export", configuration) { }


        /// <summary>
        /// Creates a new serializer using the given <see cref="ExcelPackage"/> and <see cref="Configuration"/>.
        /// <remarks>
        /// The <paramref name="package"/> will <b><i>not</i></b> be disposed of when the serializer is disposed.
        /// The package will <b><i>not</i></b> be saved by the serializer.
        /// A new worksheet will be added to the package.
        /// </remarks>
        /// </summary>
        /// <param name="package">The package to write the data to.</param>
        /// <param name="sheetName">The name of the sheet to write to.</param>
        /// <param name="configuration">The configuration.</param>
        public ExcelSerializer(ExcelPackage package, string sheetName, CsvConfiguration configuration = null)
            : this(package, package.GetOrAddWorksheet(sheetName), configuration) { }


        /// <summary>
        /// Creates a new serializer using the given <see cref="ExcelPackage"/> and <see cref="ExcelWorksheet"/>.
        /// <remarks>
        /// The <paramref name="worksheet"/> will <b><i>not</i></b> be disposed of when the serializer is disposed.
        /// The package will <b><i>not</i></b> be saved by the serializer.
        /// </remarks>
        /// </summary>
        /// <param name="package">The package to write the data to.</param>
        /// <param name="worksheet">The worksheet to write the data to.</param>
        /// <param name="configuration">The configuration</param>
        public ExcelSerializer(ExcelPackage package, ExcelWorksheet worksheet, CsvConfiguration configuration = null)
            : this(package, (ExcelRangeBase)worksheet.Cells, configuration) { }


        /// <summary>
        /// Creates a new serializer using the given <see cref="ExcelPackage"/> and <see cref="ExcelRange"/>.
        /// </summary>
        /// <param name="package">The package to write the data to.</param>
        /// <param name="range">The range to write the data to.</param>
        /// <param name="configuration">The configuration</param>
        public ExcelSerializer(ExcelPackage package, ExcelRange range, CsvConfiguration configuration = null)
            : this(package, (ExcelRangeBase)range, configuration) { }


        private ExcelSerializer(ExcelPackage package, ExcelRangeBase range, CsvConfiguration configuration)
        {
            Package = package;
            this.range = range;
            Configuration = configuration ?? new CsvConfiguration(CultureInfo.CurrentCulture);
            Configuration.ShouldQuote = (field, ctx) => false;
            Context = new WritingContext(TextWriter.Null, Configuration, false);
        }


        /// <summary>
        /// Gets the writing context.
        /// </summary>
        public WritingContext Context { get; }

        /// <summary>
        /// Gets the configuration.
        /// </summary>
        public CsvConfiguration Configuration { get; }

        /// <summary>
        /// Gets the package to which the data is being written.
        /// </summary>
        /// <value>
        /// The package.
        /// </value>
        public ExcelPackage Package { get; }

        /// <summary>
        /// Gets and sets the number of rows to offset the start position from.
        /// </summary>
        public int RowOffset { get; set; } = 0;

        /// <summary>
        /// Gets and sets the number of columns to offset the start position from.
        /// </summary>
        public int ColumnOffset { get; set; } = 0;


        /// <summary>
        /// Writes a record to the Excel file.
        /// </summary>
        /// <param name="record">The record to write.</param>
        /// <exception cref="ObjectDisposedException">
        /// Thrown is the serializer has been disposed.
        /// </exception>
        public virtual void Write(string[] record)
        {
            CheckDisposed();

            for (var i = 0; i < record.Length; i++) {
                var row = range.Start.Row + currentRow + RowOffset - 1;
                var column = range.Start.Column + ColumnOffset + i;
                range.Worksheet.SetValue(row, column, ReplaceHexadecimalSymbols(record[i]));
            }

            currentRow++;
        }


        /// <summary>
        /// Writes asynchronously a record to the Excel file.
        /// </summary>
        /// <param name="record">The record to write.</param>
        /// <returns></returns>
        public Task WriteAsync(string[] record)
        {
            Write(record);
            return Task.CompletedTask;
        }


        /// <summary>
        /// Implementation forced by CsvHelper : <see cref="IParser"/>.
        /// </summary>
        public void WriteLine() { }


        /// <summary>
        /// Implementation forced by CsvHelper : <see cref="IParser"/>
        /// </summary>
        public Task WriteLineAsync()
        {
            WriteLine();
            return Task.CompletedTask;
        }


        ISerializerConfiguration ISerializer.Configuration => Configuration;


        /// <summary>
        /// Replaces the hexadecimal symbols.
        /// </summary>
        /// <param name="text">The text to replace.</param>
        /// <returns>The input</returns>
        protected static string ReplaceHexadecimalSymbols(string text)
        {
            if (String.IsNullOrEmpty(text)) return text;
            return Regex.Replace(text, "[\x00-\x08\x0B\x0C\x0E-\x1F]", String.Empty, RegexOptions.Compiled);
        }


        /// <summary>
        /// Performs application-defined tasks associated with freeing, releasing, or resetting unmanaged resources.
        /// </summary>
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }


        public ValueTask DisposeAsync()
        {
            Dispose();
            return default;
        }


        /// <summary>
        /// Finalizes an instance of the <see cref="ExcelSerializer"/> class.
        /// </summary>
        ~ExcelSerializer()
        {
            Dispose(false);
        }


        /// <summary>
        /// Releases unmanaged and - optionally - managed resources.
        /// </summary>
        /// <param name="disposing"><c>true</c> to release both managed and unmanaged resources; <c>false</c> to release only unmanaged resources.</param>
        protected virtual void Dispose(bool disposing)
        {
            if (disposed) return;
            if (disposing) {
                if (disposePackage) {
                    Package?.SaveAs(new FileInfo(path));
                    Package?.Dispose();
                }
            }

            disposed = true;
        }


        /// <summary>
        /// Checks if the instance has been disposed of.
        /// </summary>
        /// <exception cref="ObjectDisposedException">
        /// Thrown is the serializer has been disposed.
        /// </exception>
        protected virtual void CheckDisposed()
        {
            if (disposed) {
                throw new ObjectDisposedException(GetType().ToString());
            }
        }
    }

}
