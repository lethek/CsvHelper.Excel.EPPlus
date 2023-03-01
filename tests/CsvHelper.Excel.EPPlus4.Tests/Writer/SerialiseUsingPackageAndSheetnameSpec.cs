namespace CsvHelper.Excel.EPPlus.Tests.Writer;

public class SerialiseUsingPackageAndSheetnameSpec : ExcelWriterTests
{
    public SerialiseUsingPackageAndSheetnameSpec() : base("serialise_by_package_and_sheetname", "a_different_sheet_name") {
        using var excelWriter = new ExcelWriter(Package, WorksheetName, leaveOpen: true);
        Run(excelWriter);
    }
}