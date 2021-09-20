namespace CsvHelper.Excel.EPPlus.Tests.Parser
{
    public class ParseUsingPathAndSheetNameSpec : ExcelParserTests
    {
        public ParseUsingPathAndSheetNameSpec() : base("parse_by_path_and_sheetname.xlsx", "a_different_sheet_name") {
            using var parser = new ExcelParser(Path, WorksheetName);
            Run(parser);
        }
    }
}
