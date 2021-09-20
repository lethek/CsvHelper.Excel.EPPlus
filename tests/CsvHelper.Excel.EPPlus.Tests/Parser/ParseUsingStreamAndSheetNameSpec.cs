using System.IO;


namespace CsvHelper.Excel.EPPlus.Tests.Parser
{
    public class ParseUsingStreamAndSheetNameSpec : ExcelParserTests
    {
        public ParseUsingStreamAndSheetNameSpec() : base("parse_by_stream_and_sheetname.xlsx", "a_different_sheet_name") {
            using var stream = File.OpenRead(Path);
            using var parser = new ExcelParser(stream, WorksheetName);
            Run(parser);
        }
    }
}
