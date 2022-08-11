namespace CsvHelper.Excel.EPPlus.Tests.Writer;

public class SerialiseUsingPathSpec : ExcelWriterTests
{
    public SerialiseUsingPathSpec() : base("serialise_by_path.xlsx") {
        using var excelWriter = new ExcelWriter(Path);
        Run(excelWriter);
    }
}