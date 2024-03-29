﻿using OfficeOpenXml;


namespace CsvHelper.Excel.EPPlus.Tests.Writer;

public class SerialiseUsingStreamAndSheetnameSpec : ExcelWriterTests
{
    public SerialiseUsingStreamAndSheetnameSpec() : base("serialise_by_workbook_and_sheetname", "a_different_sheet_name") {
        _stream = new MemoryStream();
        using var excelWriter = new ExcelWriter(_stream, WorksheetName, leaveOpen: true);
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