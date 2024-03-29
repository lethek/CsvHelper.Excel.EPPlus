﻿using CsvHelper.Excel.EPPlus.Tests.Common;

using FluentAssertions;

using OfficeOpenXml;


namespace CsvHelper.Excel.EPPlus.Tests.Parser;

public abstract class ExcelParserTests : IDisposable
{
    protected readonly Person[] Values = {
        new() { Id = null, Name = "Bill", Age = 40, Empty = "" },
        new() { Id = 5, Name = "Ben", Age = 30, Empty = "" },
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

    protected ExcelParserTests(string path, string worksheetName = "Export", int startRow = 1, int startColumn = 1) {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

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
        reader.Context.AutoMap<Person>();
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