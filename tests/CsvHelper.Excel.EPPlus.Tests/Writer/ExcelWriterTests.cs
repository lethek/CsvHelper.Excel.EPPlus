using System;
using System.IO;
using CsvHelper.Excel.EPPlus.Tests.Common;
using FluentAssertions;

using OfficeOpenXml;

using Xunit;


namespace CsvHelper.Excel.EPPlus.Tests.Writer
{
    public abstract class ExcelWriterTests : IDisposable
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


        protected ExcelWriterTests(string path, string worksheetName = "Export", int startRow = 1, int startColumn = 1) {
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
}
