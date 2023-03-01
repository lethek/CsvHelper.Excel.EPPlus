namespace CsvHelper.Excel.EPPlus.Tests.Parser;

public class ParseUsingWorksheetSpec : ExcelParserTests
{
    public ParseUsingWorksheetSpec() : base("parse_by_worksheet") {
        using var parser = new ExcelParser(Worksheet);
        Run(parser);
    }
}