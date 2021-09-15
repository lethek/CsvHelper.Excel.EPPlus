using System;
using System.IO;

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
                new() { Id = null, Name = "Ben", Age = 20, Empty = null },
                new() { Id = null, Name = "Weed", Age = 30, Empty = "" }
            };

            private ExcelPackage _package;
            private ExcelWorksheet _worksheet;

            protected abstract string Path { get; }

            protected virtual string WorksheetName => "Export";

            protected virtual int StartRow => 1;

            protected virtual int StartColumn => 1;

            protected ExcelPackage Package => _package ??= Helpers.GetOrCreatePackage(Path, WorksheetName);

            protected ExcelWorksheet Worksheet => _worksheet ??= Package.GetOrAddWorksheet(WorksheetName);


            protected void Run(ExcelSerializer serialiser)
            {
                using var writer = new CsvWriter(serialiser);
                writer.Configuration.AutoMap<Person>();
                writer.WriteRecords(Values);
            }


            [Fact]
            public void TheFileIsAValidExcelFile()
            {
                Assert.NotNull(Package);
            }


            [Fact]
            public void TheExcelWorkbookHeadersAreCorrect()
            {
                int column = StartColumn;
                Assert.Equal(nameof(Person.Id), Worksheet.GetValue(StartRow, column++));
                Assert.Equal(nameof(Person.Name), Worksheet.GetValue(StartRow, column++));
                Assert.Equal(nameof(Person.Age), Worksheet.GetValue(StartRow, column++));
                Assert.Equal(nameof(Person.Empty), Worksheet.GetValue(StartRow, column++));
            }


            [Fact]
            public void TheExcelWorkbookValuesAreCorrect()
            {
                for (int i = 0; i < Values.Length; i++) {
                    int column = StartColumn;
                    Assert.Equal(Values[i].Id, Worksheet.GetValue<int?>(StartRow + i + 1, column++));
                    Assert.Equal(Values[i].Name, Worksheet.GetValue(StartRow + i + 1, column++));
                    Assert.Equal(Values[i].Age, Worksheet.GetValue<int>(StartRow + i + 1, column++));
                    Assert.Equal(Values[i].Empty ?? "", Worksheet.GetValue<string>(StartRow + i + 1, column++));
                }
            }


            protected virtual void Dispose(bool disposing)
            {
                if (disposing) {
                    _package?.Save();
                    _package?.Dispose();
                    _worksheet?.Dispose();
                    File.Delete(Path);
                }
            }


            public void Dispose()
            {
                Dispose(true);
                GC.SuppressFinalize(this);
            }
        }


        public class SerialiseUsingPathSpec : Spec
        {
            public SerialiseUsingPathSpec()
            {
                using var serialiser = new ExcelSerializer(Path);
                Run(serialiser);
            }


            protected sealed override string Path => "serialise_by_path.xlsx";
        }


        public class SerialiseUsingPathWithOffsetsSpec : Spec
        {
            public SerialiseUsingPathWithOffsetsSpec()
            {
                using var serialiser = new ExcelSerializer(Path) { ColumnOffset = StartColumn - 1, RowOffset = StartRow - 1 };
                Run(serialiser);
            }


            protected override int StartColumn => 5;

            protected override int StartRow => 5;

            protected sealed override string Path => "serialise_by_path_with_offsets.xlsx";
        }


        public class SerialiseUsingPathAndSheetnameSpec : Spec
        {
            public SerialiseUsingPathAndSheetnameSpec()
            {
                using var serialiser = new ExcelSerializer(Path, WorksheetName);
                Run(serialiser);
            }


            protected sealed override string Path => "serialise_by_path_and_sheetname.xlsx";

            protected override string WorksheetName => "a_different_sheet_name";
        }


        public class SerialiseUsingPackageSpec : Spec
        {
            public SerialiseUsingPackageSpec()
            {
                using var serialiser = new ExcelSerializer(Package);
                Run(serialiser);
            }


            protected override string Path => "serialise_by_package.xlsx";
        }


        public class SerialiseUsingPackageAndSheetnameSpec : Spec
        {
            public SerialiseUsingPackageAndSheetnameSpec()
            {
                using var serialiser = new ExcelSerializer(Package, WorksheetName);
                Run(serialiser);
            }


            protected override string Path => "serialise_by_package_and_sheetname.xlsx";

            protected override string WorksheetName => "a_different_sheet_name";
        }


        public class SerialiseUsingWorksheetSpec : Spec
        {
            public SerialiseUsingWorksheetSpec()
            {
                using var serialiser = new ExcelSerializer(Package, Worksheet);
                Run(serialiser);
            }


            protected override string Path => "serialise_by_worksheet.xlsx";

            protected override string WorksheetName => "a_different_sheetname";
        }


        public class SerialiseUsingRangeSpec : Spec
        {
            public SerialiseUsingRangeSpec()
            {
                var range = Worksheet.Cells[StartRow, StartColumn, StartRow + Values.Length, StartColumn + 1];
                using var serialiser = new ExcelSerializer(Package, range);
                Run(serialiser);
            }

            protected override int StartRow => 4;

            protected override int StartColumn => 8;

            protected override string Path => "serialise_by_range.xlsx";
        }

    }

}
