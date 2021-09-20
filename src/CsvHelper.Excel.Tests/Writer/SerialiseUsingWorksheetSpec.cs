namespace CsvHelper.Excel.Tests.Writer
{
    public class SerialiseUsingWorksheetSpec : ExcelWriterTests
    {
        public SerialiseUsingWorksheetSpec() : base("serialise_by_worksheet.xlsx", "a_different_sheetname") {
            using var excelWriter = new ExcelWriter(Package, Worksheet);
            Run(excelWriter);
        }
    }
}
