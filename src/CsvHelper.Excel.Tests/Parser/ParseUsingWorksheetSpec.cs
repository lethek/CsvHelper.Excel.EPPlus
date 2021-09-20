namespace CsvHelper.Excel.Tests.Parser
{
    public class ParseUsingWorksheetSpec : ExcelParserTests
    {
        public ParseUsingWorksheetSpec() : base("parse_by_worksheet.xlsx") {
            using var parser = new ExcelParser(Worksheet);
            Run(parser);
        }
    }
}
