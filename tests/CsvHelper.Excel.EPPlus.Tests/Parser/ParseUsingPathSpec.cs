namespace CsvHelper.Excel.EPPlus.Tests.Parser;

public class ParseUsingPathSpec : ExcelParserTests
{
    public ParseUsingPathSpec() : base("parse_by_path") {
        using var parser = new ExcelParser(Path);
        Run(parser);
    }
}