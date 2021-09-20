using System.IO;

using OfficeOpenXml;


namespace CsvHelper.Excel.Tests.Writer
{
    public class SerialiseUsingStreamSpec : ExcelWriterTests
    {
        public SerialiseUsingStreamSpec() : base("serialise_by_workbook.xlsx") {
            _stream = new MemoryStream();
            using var excelWriter = new ExcelWriter(_stream);
            Run(excelWriter);
        }

        protected override ExcelPackage CreatePackage() {
            _stream.Position = 0;
            return new ExcelPackage(_stream);
        }

        protected override void Dispose(bool disposing) {
            base.Dispose(disposing);
            _stream.Dispose();
        }

        private readonly Stream _stream;
    }
}
