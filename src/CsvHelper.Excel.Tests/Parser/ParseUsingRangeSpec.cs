namespace CsvHelper.Excel.Tests.Parser
{
    public class ParseUsingRangeSpec : ExcelParserTests
    {
        public ParseUsingRangeSpec() : base("parse_with_range.xlsx", "Export", 4, 5) {
            var range = Worksheet.Cells[StartRow, StartColumn, StartRow + Values.Length, StartColumn + 3];
            using var parser = new ExcelParser(range);
            Run(parser);
        }
    }
}
