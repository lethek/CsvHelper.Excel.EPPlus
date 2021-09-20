using System.IO;


namespace CsvHelper.Excel.Tests.Parser
{
    public class ParseUsingStreamSpec : ExcelParserTests
    {
        public ParseUsingStreamSpec() : base("parse_by_stream.xlsx") {
            using var stream = File.OpenRead(Path);
            using var parser = new ExcelParser(stream);
            Run(parser);
        }
    }
}
