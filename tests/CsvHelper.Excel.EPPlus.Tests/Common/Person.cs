﻿namespace CsvHelper.Excel.EPPlus.Tests.Common;

public record Person
{
    public int? Id { get; init; }
    public string Name { get; init; }
    public int Age { get; init; }
    public string Empty { get; init; }
}