namespace CsvHelper.Excel.EPPlus.Tests.Writer
{
    public class SerialiseUsingPathAndSheetnameSpec : ExcelWriterTests
    {
        public SerialiseUsingPathAndSheetnameSpec() : base("serialise_by_path_and_sheetname.xlsx", "a_different_sheet_name") {
            using var excelWriter = new ExcelWriter(Path, WorksheetName);
            Run(excelWriter);
        }
    }
}
