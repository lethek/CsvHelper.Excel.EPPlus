using OfficeOpenXml;


namespace CsvHelper.Excel.Tests
{

    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;

    using Xunit;


    public class ExcelParserTests
    {
        public abstract class Spec : IDisposable
        {
            protected readonly Person[] Values = {
                new Person { Name = "Bill", Id = null, Age = 40, Empty = "" },
                new Person { Name = "Ben", Id = null, Age = 30, Empty = null },
                new Person { Name = "Weed", Id = null, Age = 40, Empty = "" }
            };

            private ExcelPackage package;
            private ExcelWorksheet worksheet;
            protected Person[] Results;


            protected Spec()
            {
                var package = Helpers.GetOrCreatePackage(Path, WorksheetName);
                var worksheet = package.GetOrAddWorksheet(WorksheetName);
                var headerRow = worksheet.Row(StartRow);
                int column = StartColumn;
                worksheet.SetValue(headerRow.Row, column++, nameof(Person.Name));
                worksheet.SetValue(headerRow.Row, column++, nameof(Person.Id));
                worksheet.SetValue(headerRow.Row, column++, nameof(Person.Age));
                worksheet.SetValue(headerRow.Row, column++, nameof(Person.Empty));
                for (int i = 0; i < Values.Length; i++) {
                    column = StartColumn;
                    var row = worksheet.Row(StartRow + i + 1);
                    worksheet.SetValue(row.Row, column++, Values[i].Name);
                    worksheet.SetValue(row.Row, column++, Values[i].Id);
                    worksheet.SetValue(row.Row, column++, Values[i].Age);
                    worksheet.SetValue(row.Row, column++, Values[i].Empty);
                }

                package.SaveAs(new FileInfo(Path));
            }


            protected abstract string Path { get; }

            protected virtual string WorksheetName => "Export";

            protected virtual int StartRow => 1;

            protected virtual int StartColumn => 1;

            protected ExcelPackage Package => package ?? (package = Helpers.GetOrCreatePackage(Path, WorksheetName));

            protected ExcelWorksheet Worksheet => worksheet ?? (worksheet = Package.GetOrAddWorksheet(WorksheetName));


            protected void Run(ExcelParser parser)
            {
                using (var reader = new CsvReader(parser)) {
                    reader.Configuration.AutoMap<Person>();
                    Results = reader.GetRecords<Person>().ToArray();
                }
            }


            [Fact]
            public void TheResultsAreNotNull()
            {
                Assert.NotNull(Results);
            }


            [Fact]
            public void TheResultsAreCorrect()
            {
                Assert.Equal(Values, Results, EqualityComparer<Person>.Default);
            }


            public void Dispose()
            {
                Package?.Dispose();
                File.Delete(Path);
            }
        }


        public class ParseUsingPathSpec : Spec
        {
            public ParseUsingPathSpec()
            {
                using (var parser = new ExcelParser(Path)) {
                    Run(parser);
                }
            }


            protected override string Path => "parse_by_path.xlsx";
        }


        public class ParseUsingPathWithOffsetsSpec : Spec
        {
            public ParseUsingPathWithOffsetsSpec()
            {
                using (var parser = new ExcelParser(Path) { ColumnOffset = StartColumn - 1, RowOffset = StartRow - 1 }) {
                    Run(parser);
                }
            }


            protected override int StartColumn => 5;

            protected override int StartRow => 5;

            protected override string Path => "parse_by_path_with_offset.xlsx";
        }


        public class ParseUsingPathAndSheetNameSpec : Spec
        {
            public ParseUsingPathAndSheetNameSpec()
            {
                using (var parser = new ExcelParser(Path, WorksheetName)) {
                    Run(parser);
                }
            }


            protected override string Path => "parse_by_path_and_sheetname.xlsx";

            protected override string WorksheetName => "a_different_sheet_name";
        }


        public class ParseUsingPackageSpec : Spec
        {
            public ParseUsingPackageSpec()
            {
                using (var parser = new ExcelParser(Package)) {
                    Run(parser);
                }
            }


            protected override string Path => "parse_by_package.xlsx";
        }


        public class ParseUsingPackageAndSheetNameSpec : Spec
        {
            public ParseUsingPackageAndSheetNameSpec()
            {
                using (var parser = new ExcelParser(Package, WorksheetName)) {
                    Run(parser);
                }
            }


            protected override string Path => "parse_by_package_and_sheetname.xlsx";

            protected override string WorksheetName => "a_different_sheet_name";
        }


        public class ParseUsingWorksheetSpec : Spec
        {
            public ParseUsingWorksheetSpec()
            {
                using (var parser = new ExcelParser(Worksheet)) {
                    Run(parser);
                }
            }


            protected override string Path => "parse_by_worksheet.xlsx";
        }


        public class ParseUsingRangeSpec : Spec
        {
            public ParseUsingRangeSpec()
            {
                var range = Worksheet.Cells[StartRow, StartColumn, StartRow + Values.Length, StartColumn + 1];
                using (var parser = new ExcelParser(range)) {
                    Run(parser);
                }
            }


            protected override int StartColumn => 5;

            protected override int StartRow => 4;

            protected override string Path => "parse_with_range.xlsx";
        }


        public class ParseWithFormulaSpec : Spec
        {
            public ParseWithFormulaSpec()
            {
                for (int i = 0; i < Values.Length; i++) {
                    var row = Worksheet.Row(2 + i);
                    Worksheet.Cells[row.Row, 3].FormulaR1C1 = $"=LEN({Worksheet.Cells[row.Row, 1].Address})*10";
                }

                Package.SaveAs(new FileInfo(Path));
                using (var parser = new ExcelParser(Path)) {
                    Run(parser);
                }
            }


            protected override string Path => "parse_with_formula.xlsx";
        }

    }

}
