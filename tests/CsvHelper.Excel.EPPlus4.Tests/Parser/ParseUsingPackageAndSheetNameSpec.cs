namespace CsvHelper.Excel.EPPlus.Tests.Parser;

public class ParseUsingPackageAndSheetNameSpec : ExcelParserTests
{
    public ParseUsingPackageAndSheetNameSpec() : base("parse_by_package_and_sheetname", "a_different_sheet_name") {
        using var parser = new ExcelParser(Package, WorksheetName);
        Run(parser);
    }
}