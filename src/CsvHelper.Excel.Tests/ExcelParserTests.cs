using System;
using System.IO;
using System.Linq;

using FluentAssertions;

using OfficeOpenXml;

using Xunit;


namespace CsvHelper.Excel.Tests
{
    public class ExcelParserTests
    {
        public abstract class Spec : IDisposable
        {
            protected readonly Person[] Values = {
                new() { Id = null, Name = "Bill", Age = 40, Empty = "" },
                new() { Id = null, Name = "Ben", Age = 30, Empty = "" },
                new() { Id = null, Name = "Weed", Age = 40, Empty = "" }
            };

            protected Person[] Results;
            protected string Path { get; }
            protected string Dir { get; }
            protected string WorksheetName { get; }
            protected int StartRow { get; }
            protected int StartColumn { get; }
            protected ExcelPackage Package { get; }
            protected ExcelWorksheet Worksheet { get; }

            protected Spec(string path, string worksheetName = "Export", int startRow = 1, int startColumn = 1) {
                Path = System.IO.Path.GetFullPath(System.IO.Path.Combine("data", Guid.NewGuid().ToString(), $"{path}.xlsx"));
                Dir = System.IO.Path.GetDirectoryName(Path);

                if (!Directory.Exists(Dir)) {
                    Directory.CreateDirectory(Dir);
                }

                WorksheetName = worksheetName;
                StartRow = startRow;
                StartColumn = startColumn;

                Package = Helpers.GetOrCreatePackage(Path, WorksheetName);
                Worksheet = Package.GetOrAddWorksheet(WorksheetName);

                var headerRow = Worksheet.Row(StartRow);
                int column = StartColumn;
                Worksheet.SetValue(headerRow.Row, column++, nameof(Person.Id));
                Worksheet.SetValue(headerRow.Row, column++, nameof(Person.Name));
                Worksheet.SetValue(headerRow.Row, column++, nameof(Person.Age));
                Worksheet.SetValue(headerRow.Row, column++, nameof(Person.Empty));
                for (int i = 0; i < Values.Length; i++) {
                    column = StartColumn;
                    var row = Worksheet.Row(StartRow + i + 1);
                    Worksheet.SetValue(row.Row, column++, Values[i].Id);
                    Worksheet.SetValue(row.Row, column++, Values[i].Name);
                    Worksheet.SetValue(row.Row, column++, Values[i].Age);
                    Worksheet.SetValue(row.Row, column++, Values[i].Empty);
                }

                Package.SaveAs(new FileInfo(Path));
            }


            protected void Run(ExcelParser parser) {
                using var reader = new CsvReader(parser);
                reader.Configuration.AutoMap<Person>();
                Results = reader.GetRecords<Person>().ToArray();
            }


            [Fact]
            public void TheResultsAreNotNull()
                => Results.Should().NotBeNull();


            [Fact]
            public void TheResultsAreCorrect()
               => Values.Should().BeEquivalentTo(Results, options => options.IncludingProperties());


            protected virtual void Dispose(bool disposing) {
                if (disposing) {
                    Package?.Dispose();
                    Worksheet?.Dispose();
                    Helpers.Delete(Path);
                }
            }


            public void Dispose() {
                Dispose(true);
                GC.SuppressFinalize(this);
            }
        }


        public class ParseUsingPathSpec : Spec
        {
            public ParseUsingPathSpec() : base("parse_by_path.xlsx") {
                using var parser = new ExcelParser(Path);
                Run(parser);
            }
        }


        public class ParseUsingPathWithOffsetsSpec : Spec
        {
            public ParseUsingPathWithOffsetsSpec(): base("parse_by_path_with_offset.xlsx", "Export", 5, 5) {
                using var parser = new ExcelParser(Path) { ColumnOffset = StartColumn - 1, RowOffset = StartRow - 1 };
                Run(parser);
            }
        }


        public class ParseUsingPathAndSheetNameSpec : Spec
        {
            public ParseUsingPathAndSheetNameSpec() : base("parse_by_path_and_sheetname.xlsx", "a_different_sheet_name") {
                using var parser = new ExcelParser(Path, WorksheetName);
                Run(parser);
            }
        }


        public class ParseUsingPackageSpec : Spec
        {
            public ParseUsingPackageSpec() : base("parse_by_package.xlsx") {
                using var parser = new ExcelParser(Package);
                Run(parser);
            }
        }


        public class ParseUsingPackageAndSheetNameSpec : Spec
        {
            public ParseUsingPackageAndSheetNameSpec() : base("parse_by_package_and_sheetname.xlsx", "a_different_sheet_name") {
                using var parser = new ExcelParser(Package, WorksheetName);
                Run(parser);
            }
        }


        public class ParseUsingWorksheetSpec : Spec
        {
            public ParseUsingWorksheetSpec() : base("parse_by_worksheet.xlsx") {
                using var parser = new ExcelParser(Worksheet);
                Run(parser);
            }
        }


        public class ParseUsingRangeSpec : Spec
        {
            public ParseUsingRangeSpec() : base("parse_with_range.xlsx", "Export", 4, 5) {
                var range = Worksheet.Cells[StartRow, StartColumn, StartRow + Values.Length, StartColumn + 1];
                using var parser = new ExcelParser(range);
                Run(parser);
            }
        }


        public class ParseWithFormulaSpec : Spec
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
}
