using System;
using System.IO;

using FluentAssertions;

using OfficeOpenXml;

using Xunit;


namespace CsvHelper.Excel.Tests
{
    public class ExcelSerializerTests
    {
        public abstract class Spec : IDisposable
        {
            protected readonly Person[] Values = {
                new() { Id = null, Name = "Bill", Age = 20, Empty = "" },
                new() { Id = null, Name = "Ben", Age = 20, Empty = "" },
                new() { Id = null, Name = "Weed", Age = 30, Empty = "" }
            };

            protected string Path { get; }
            protected string Dir { get; }
            protected string WorksheetName { get; }
            protected int StartRow { get; }
            protected int StartColumn { get; }

            protected ExcelPackage Package => _package ??= CreatePackage();
            protected ExcelWorksheet Worksheet => _worksheet ??= CreateWorksheet();

            private ExcelPackage _package;
            private ExcelWorksheet _worksheet;


            protected Spec(string path, string worksheetName = "Export", int startRow = 1, int startColumn = 1) {
                Path = System.IO.Path.GetFullPath(System.IO.Path.Combine("data", Guid.NewGuid().ToString(), $"{path}.xlsx"));

                Dir = System.IO.Path.GetDirectoryName(Path);
                if (!Directory.Exists(Dir)) {
                    Directory.CreateDirectory(Dir);
                }

                WorksheetName = worksheetName;
                StartRow = startRow;
                StartColumn = startColumn;
            }


            protected virtual ExcelPackage CreatePackage()
                => Helpers.GetOrCreatePackage(Path, WorksheetName);


            protected virtual ExcelWorksheet CreateWorksheet()
                => Package.GetOrAddWorksheet(WorksheetName);


            protected void Run(ExcelWriter excelWriter) {
                excelWriter.Context.AutoMap<Person>();
                excelWriter.WriteRecords(Values);
            }


            [Fact]
            public void TheFileIsAValidExcelFile()
                => Package.Should().NotBeNull();


            [Fact]
            public void TheExcelWorkbookHeadersAreCorrect() {
                int column = StartColumn;
                nameof(Person.Id).Should().Be(Worksheet.GetValue(StartRow, column++).ToString());
                nameof(Person.Name).Should().Be(Worksheet.GetValue(StartRow, column++).ToString());
                nameof(Person.Age).Should().Be(Worksheet.GetValue(StartRow, column++).ToString());
                nameof(Person.Empty).Should().Be(Worksheet.GetValue(StartRow, column++).ToString());
            }


            [Fact]
            public void TheExcelWorkbookValuesAreCorrect() {
                for (int i = 0; i < Values.Length; i++) {
                    int column = StartColumn;
                    Values[i].Id.Should().Be(Worksheet.GetValue<int?>(StartRow + i + 1, column++).As<int?>());
                    Values[i].Name.Should().Be(Worksheet.GetValue(StartRow + i + 1, column++).As<string>());
                    Values[i].Age.Should().Be(Worksheet.GetValue<int>(StartRow + i + 1, column++).As<int>());
                    Values[i].Empty.Should().Be(Worksheet.GetValue<string>(StartRow + i + 1, column++).As<string>());
                }
            }


            protected virtual void Dispose(bool disposing) {
                if (disposing) {
                    _package?.Save();
                    _package?.Dispose();
                    _worksheet?.Dispose();
                    Helpers.Delete(Path);
                }
            }


            public void Dispose() {
                Dispose(true);
                GC.SuppressFinalize(this);
            }
        }


        public class SerialiseUsingPathSpec : Spec
        {
            public SerialiseUsingPathSpec() : base("serialise_by_path.xlsx") {
                using var excelWriter = new ExcelWriter(Path);
                Run(excelWriter);
            }
        }


        public class SerialiseUsingPathWithOffsetsSpec : Spec
        {
            public SerialiseUsingPathWithOffsetsSpec() : base("serialise_by_path_with_offsets.xlsx", "Export", 5, 5) {
                using var excelWriter = new ExcelWriter(Path) {
                    ColumnOffset = StartColumn - 1,
                    RowOffset = StartRow - 1
                };
                Run(excelWriter);
            }
        }


        public class SerialiseUsingPathAndSheetnameSpec : Spec
        {
            public SerialiseUsingPathAndSheetnameSpec() : base("serialise_by_path_and_sheetname.xlsx", "a_different_sheet_name") {
                using var excelWriter = new ExcelWriter(Path, WorksheetName);
                Run(excelWriter);
            }
        }


        public class SerialiseUsingPackageSpec : Spec
        {
            public SerialiseUsingPackageSpec() : base("serialise_by_package.xlsx") {
                using var excelWriter = new ExcelWriter(Package);
                Run(excelWriter);
            }
        }


        public class SerialiseUsingPackageAndSheetnameSpec : Spec
        {
            public SerialiseUsingPackageAndSheetnameSpec() : base("serialise_by_package_and_sheetname.xlsx", "a_different_sheet_name") {
                using var excelWriter = new ExcelWriter(Package, WorksheetName);
                Run(excelWriter);
            }
        }


        public class SerialiseUsingWorksheetSpec : Spec
        {
            public SerialiseUsingWorksheetSpec() : base("serialise_by_worksheet.xlsx", "a_different_sheetname") {
                using var excelWriter = new ExcelWriter(Package, Worksheet);
                Run(excelWriter);
            }
        }


        public class SerialiseUsingRangeSpec : Spec
        {
            public SerialiseUsingRangeSpec() : base("serialise_by_range.xlsx", "Export", 4, 8) {
                var range = Worksheet.Cells[StartRow, StartColumn, StartRow + Values.Length, StartColumn + 1];
                using var excelWriter = new ExcelWriter(Package, range);
                Run(excelWriter);
            }
        }


        public class SerialiseUsingStreamSpec : Spec
        {
            public SerialiseUsingStreamSpec() : base("serialise_by_workbook.xlsx") {
                _stream = new MemoryStream();
                using var excelWriter = new ExcelWriter(_stream);
                Run(excelWriter);
            }

            protected override ExcelPackage CreatePackage() {
                _stream.Position = 0;
                return new ExcelPackage(_stream);
            }

            protected override void Dispose(bool disposing) {
                base.Dispose(disposing);
                _stream.Dispose();
            }

            private readonly Stream _stream;
        }


        public class SerialiseUsingStreamAndSheetnameSpec : Spec
        {
            public SerialiseUsingStreamAndSheetnameSpec() : base("serialise_by_workbook_and_sheetname.xlsx", "a_different_sheet_name") {
                _stream = new MemoryStream();
                using var excelWriter = new ExcelWriter(_stream, WorksheetName);
                Run(excelWriter);
            }

            protected override ExcelPackage CreatePackage() {
                _stream.Position = 0;
                return new ExcelPackage(_stream);
            }

            protected override void Dispose(bool disposing) {
                base.Dispose(disposing);
                _stream.Dispose();
            }

            private readonly Stream _stream;
        }
    }
}
