namespace CsvHelper.Excel.EPPlus.Tests.Parser
{
    public class ParseUsingPackageSpec : ExcelParserTests
    {
        public ParseUsingPackageSpec() : base("parse_by_package.xlsx") {
            using var parser = new ExcelParser(Package);
            Run(parser);
        }
    }
}
