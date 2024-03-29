﻿namespace CsvHelper.Excel.EPPlus.Tests.Writer;

public class SerialiseUsingRangeSpec : ExcelWriterTests
{
    public SerialiseUsingRangeSpec() : base("serialise_by_range", "Export", 4, 8) {
        var range = Worksheet.Cells[StartRow, StartColumn, StartRow + Values.Length, StartColumn + 1];
        using var excelWriter = new ExcelWriter(Package, range, leaveOpen: true);
        Run(excelWriter);
    }
}