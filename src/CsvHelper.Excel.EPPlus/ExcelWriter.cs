using System.Globalization;
using System.Runtime.CompilerServices;

using CsvHelper.Configuration;

using OfficeOpenXml;


namespace CsvHelper.Excel.EPPlus;

/// <summary>
/// Defines methods used to serialize data into an Excel (2007+) file.
/// </summary>
public class ExcelWriter : CsvWriter
{
    /// <summary>
    /// Gets the package to which the data is being written.
    /// </summary>
    /// <value>The package.</value>
    public ExcelPackage Package { get; }


    /// <summary>
    /// Gets and sets the number of rows to offset the start position from.
    /// </summary>
    public int RowOffset { get; set; }


    /// <summary>
    /// Gets and sets the number of columns to offset the start position from.
    /// </summary>
    public int ColumnOffset { get; set; }


    /// <summary>
    /// Creates a new serializer using a new <see cref="ExcelPackage"/> saved to the given <paramref name="path"/>.
    /// <remarks>
    /// The package will not be saved until the serializer is disposed.
    /// </remarks>
    /// </summary>
    /// <param name="path">The path to which to save the package.</param>
    /// <param name="sheetName">The name of the sheet to which to save.</param>
    public ExcelWriter(string path, string sheetName = "Export", CsvConfiguration configuration = null)
        : this(new ExcelPackage(new FileInfo(path)), sheetName, configuration) {
        _isPackageOwner = true;
    }


    /// <summary>
    /// Creates a new serializer using a new <see cref="ExcelPackage"/> saved to the given <paramref name="stream"/>.
    /// <remarks>
    /// The package will not be saved until the serializer is disposed.
    /// </remarks>
    /// </summary>
    /// <param name="stream">The stream to which to save the package.</param>
    /// <param name="sheetName">The name of the sheet to which to save</param>
    public ExcelWriter(Stream stream, string sheetName = "Export", CsvConfiguration configuration = null)
        : this(new ExcelPackage(), sheetName, configuration) {
        _stream = stream;
        _isPackageOwner = true;
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
    /// <param name="sheetName">The name of the sheet to write to.</param>
    /// <param name="configuration">The configuration.</param>
    public ExcelWriter(ExcelPackage package, string sheetName = "Export", CsvConfiguration configuration = null)
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
    public ExcelWriter(ExcelPackage package, ExcelWorksheet worksheet, CsvConfiguration configuration = null)
        : this(package, worksheet.Cells, configuration) { }


    /// <summary>
    /// Creates a new serializer using the given <see cref="ExcelPackage"/> and <see cref="ExcelRange"/>.
    /// </summary>
    /// <param name="package">The package to write the data to.</param>
    /// <param name="range">The range to write the data to.</param>
    /// <param name="configuration">The configuration</param>
    public ExcelWriter(ExcelPackage package, ExcelRange range, CsvConfiguration configuration = null)
        : this(package, (ExcelRangeBase)range, configuration ?? new CsvConfiguration(CultureInfo.InvariantCulture) { LeaveOpen = true }) { }


    private ExcelWriter(ExcelPackage package, ExcelRangeBase range, CsvConfiguration configuration)
        : base(TextWriter.Null, configuration) {
        configuration.Validate();

        Package = package;
        _range = range;

        _leaveOpen = configuration.LeaveOpen;
        _sanitizeForInjection = configuration.InjectionOptions != InjectionOptions.None;
    }


    /// <inheritdoc/>
    public override void WriteField(string field, bool shouldQuote) {
        if (_sanitizeForInjection) {
            field = SanitizeForInjection(field);
        }

        WriteToCell(field);
        _index++;
    }


    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    private void WriteToCell(string value) {
        _range.Worksheet.SetValue(
            _range.Start.Row + RowOffset + _row - 1,
            _range.Start.Column + ColumnOffset + _index - 1,
            value
        );
    }


    /// <inheritdoc/>
    public override void NextRecord() {
        Flush();
        _index = 1;
        _row++;
    }


    /// <inheritdoc/>
    public override async Task NextRecordAsync() {
        await FlushAsync().ConfigureAwait(false);
        _index = 1;
        _row++;
    }


    /// <summary>
    /// Implementation forced by CsvHelper : <see cref="IParser"/>.
    /// </summary>
    public override void Flush() { }


    /// <summary>
    /// Implementation forced by CsvHelper : <see cref="IParser"/>.
    /// </summary>
    public override Task FlushAsync()
        => Task.CompletedTask;


    /// <summary>
    /// Implementation forced by CsvHelper : <see cref="IParser"/>.
    /// </summary>
    public void WriteLine() { }


    /// <summary>
    /// Implementation forced by CsvHelper : <see cref="IParser"/>
    /// </summary>
    public Task WriteLineAsync()
        => Task.CompletedTask;


    /// <summary>
    /// Releases unmanaged and - optionally - managed resources.
    /// </summary>
    /// <param name="disposing"><c>true</c> to release both managed and unmanaged resources; <c>false</c> to release only unmanaged resources.</param>
    protected override void Dispose(bool disposing) {
        if (_disposed) {
            return;
        }

        Flush();
        if (_stream != null) {
            Package.SaveAs(_stream);
            _stream.Flush();
        } else {
            Package.Save();
        }

        if (disposing) {
            if (!_leaveOpen || _isPackageOwner) {
                Package.Dispose();
            }
            if (!_leaveOpen) {
                _stream?.Dispose();
            }
        }
        _disposed = true;
    }


#if NETSTANDARD2_1_OR_GREATER
        /// <inheritdoc/>
        protected override async ValueTask DisposeAsync(bool disposing) {
            if (_disposed) {
                return;
            }

            await FlushAsync().ConfigureAwait(false);
            if (_stream != null) {
                Package.SaveAs(_stream);
                await _stream.FlushAsync().ConfigureAwait(false);
            } else {
                Package.Save();
            }

            if (disposing) {
                //Dispose managed state (managed objects)
                if (!_leaveOpen || _isPackageOwner) {
                    Package.Dispose();
                }
                if (!_leaveOpen) {
                    if (_stream != null) {
                        await _stream.DisposeAsync().ConfigureAwait(false);
                    }
                }
            }

            // Free unmanaged resources (unmanaged objects) and override finalizer
            // Set large fields to null
            _disposed = true;
        }
#endif


    private readonly bool _isPackageOwner;
    private readonly bool _leaveOpen;
    private readonly bool _sanitizeForInjection;

    private readonly Stream _stream;
    private readonly ExcelRangeBase _range;

    private bool _disposed;

    private int _row = 1;
    private int _index = 1;
}