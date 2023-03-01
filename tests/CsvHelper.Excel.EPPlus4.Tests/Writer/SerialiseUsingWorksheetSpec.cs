namespace CsvHelper.Excel.EPPlus.Tests.Writer;

public class SerialiseUsingWorksheetSpec : ExcelWriterTests
{
    public SerialiseUsingWorksheetSpec() : base("serialise_by_worksheet", "a_different_sheetname") {
        using var excelWriter = new ExcelWriter(Package, Worksheet, leaveOpen: true);
        Run(excelWriter);
    }
}