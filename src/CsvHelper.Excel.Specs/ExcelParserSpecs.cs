using OfficeOpenXml;


namespace CsvHelper.Excel.Specs
{

    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;

    using Xunit;


    public class ExcelParserSpecs
    {
        public abstract class Spec : IDisposable
        {
            protected readonly Person[] Values = {
                new Person { Name = "Bill", Age = 40 },
                new Person { Name = "Ben", Age = 30 },
                new Person { Name = "Weed", Age = 40 }
            };

            private ExcelPackage package;
            private ExcelWorksheet worksheet;
            protected Person[] Results;


            protected Spec()
            {
                var package = Helpers.GetOrCreatePackage(Path, WorksheetName);
                var worksheet = package.GetOrAddWorksheet(WorksheetName);
                var headerRow = worksheet.Row(StartRow);
                worksheet.SetValue(headerRow.Row, StartColumn, nameof(Person.Name));
                worksheet.SetValue(headerRow.Row, StartColumn + 1, nameof(Person.Age));
                for (int i = 0; i < Values.Length; i++) {
                    var row = worksheet.Row(StartRow + i + 1);
                    worksheet.SetValue(row.Row, StartColumn, Values[i].Name);
                    worksheet.SetValue(row.Row, StartColumn + 1, Values[i].Age);
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
                    Worksheet.Cells[row.Row, 2].FormulaR1C1 = $"=LEN({Worksheet.Cells[row.Row, 1].Address})*10";
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
