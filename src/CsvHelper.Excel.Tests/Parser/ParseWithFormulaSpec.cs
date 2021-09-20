using System.IO;


namespace CsvHelper.Excel.Tests.Parser
{
    public class ParseWithFormulaSpec : ExcelParserTests
    {
        public ParseWithFormulaSpec() : base("parse_with_formula.xlsx") {
            for (int i = 0; i < Values.Length; i++) {
                var row = Worksheet.Row(2 + i);
                Worksheet.Cells[row.Row, 3].FormulaR1C1 = $"=LEN({Worksheet.Cells[row.Row, 2].Address})*10";
            }
            Package.SaveAs(new FileInfo(Path));
            using var parser = new ExcelParser(Path);
            Run(parser);
        }
    }
}
