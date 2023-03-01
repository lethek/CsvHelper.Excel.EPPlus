namespace CsvHelper.Excel.EPPlus.Tests.Writer;

public class SerialiseUsingPathSpec : ExcelWriterTests
{
    public SerialiseUsingPathSpec() : base("serialise_by_path") {
        using var excelWriter = new ExcelWriter(Path);
        Run(excelWriter);
    }
}