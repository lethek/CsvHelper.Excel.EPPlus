﻿using OfficeOpenXml;


namespace CsvHelper.Excel.EPPlus.Tests.Writer;

public class SerialiseUsingStreamSpec : ExcelWriterTests
{
    public SerialiseUsingStreamSpec() : base("serialise_by_workbook") {
        _stream = new MemoryStream();
        using var excelWriter = new ExcelWriter(_stream, leaveOpen: true);
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