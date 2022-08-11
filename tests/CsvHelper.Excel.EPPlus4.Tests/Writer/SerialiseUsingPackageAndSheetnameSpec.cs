namespace CsvHelper.Excel.EPPlus.Tests.Writer;

public class SerialiseUsingPackageAndSheetnameSpec : ExcelWriterTests
{
    public SerialiseUsingPackageAndSheetnameSpec() : base("serialise_by_package_and_sheetname.xlsx", "a_different_sheet_name") {
        using var excelWriter = new ExcelWriter(Package, WorksheetName);
        Run(excelWriter);
    }
}