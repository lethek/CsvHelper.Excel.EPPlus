using OfficeOpenXml;

namespace CsvHelper.Excel.EPPlus.Tests.Writer;

public class SerialiseUsingPackageSpec : ExcelWriterTests
{
    public SerialiseUsingPackageSpec() : base("serialise_by_package") {
        using var excelWriter = new ExcelWriter(Package, leaveOpen: true);
        Run(excelWriter);
    }
}