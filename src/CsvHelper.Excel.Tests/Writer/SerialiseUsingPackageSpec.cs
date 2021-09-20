namespace CsvHelper.Excel.Tests.Writer
{
    public class SerialiseUsingPackageSpec : ExcelWriterTests
    {
        public SerialiseUsingPackageSpec() : base("serialise_by_package.xlsx") {
            using var excelWriter = new ExcelWriter(Package);
            Run(excelWriter);
        }
    }
}
