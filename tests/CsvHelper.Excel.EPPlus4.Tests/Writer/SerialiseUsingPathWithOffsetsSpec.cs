namespace CsvHelper.Excel.EPPlus.Tests.Writer;

public class SerialiseUsingPathWithOffsetsSpec : ExcelWriterTests
{
    public SerialiseUsingPathWithOffsetsSpec() : base("serialise_by_path_with_offsets", "Export", 5, 5) {
        using var excelWriter = new ExcelWriter(Path) {
            ColumnOffset = StartColumn - 1,
            RowOffset = StartRow - 1
        };
        Run(excelWriter);
    }
}