[![Build Status](https://dev.azure.com/lethek0447/lethek/_apis/build/status/lethek.CsvHelper.Excel)](https://dev.azure.com/lethek0447/lethek/_build/latest?definitionId=2)

# Csv Helper for Excel

***This project has been forked from https://github.com/christophano/CsvHelper.Excel and https://github.com/youngcm2/CsvHelper.Excel and heavily modified; primarily so that it can be used with the final LGPL version of [EPPlus](https://github.com/JanKallman/EPPlus) instead of ClosedXml, because it works with encrypted/password-protected Excel documents.***

***NuGet packages of this fork are available from MyGet:  https://www.myget.org/feed/lethek/package/nuget/CsvHelper.Excel***

CsvHelper for Excel is an extension that links two excellent libraries, [CsvHelper](https://joshclose.github.io/CsvHelper/) and [EPPlus](https://github.com/JanKallman/EPPlus).
It provides an implementation of `IParser` and `ISerializer` from [CsvHelper](https://joshclose.github.io/CsvHelper/) that read and write to Excel using [EPPlus](https://github.com/JanKallman/EPPlus).

If you need to parse or serialize a password-protected Excel document you will need to create an instance of `ExcelPackage` yourself (e.g. `new ExcelPackage("file.xlsx", password)`) and use one of the constructors described below which take that as a parameter.

## ExcelParser
`ExcelParser` implements `IParser` and allows you to specify the path of the Excel package, pass an instance of `ExcelPackage`, `ExcelWorkbook`, `ExcelWorksheet`, `ExcelRange` or a `Stream` that you have already loaded to use as the data source.

When the path is passed to the constructor then the workbook loading and disposal is handled by the parser. By default the first worksheet is used as the data source.
```csharp
using var reader = new CsvReader(new ExcelParser("path/to/file.xlsx"));
var people = reader.GetRecords<Person>();
```

When an instance of `ExcelPackage` is passed to the constructor then disposal will not be handled by the parser. By default the first worksheet is used as the data source.
```csharp
using var package = new ExcelPackage("path/to/file.xlsx");
// do stuff with the package
using var reader = new CsvReader(new ExcelParser(package));
var people = reader.GetRecords<Person>();
// do other stuff with workbook
```

When an instance of `ExcelWorksheet` is passed to the constructor then disposal will not be handled by the parser and the worksheet will be used as the data source.
```csharp
using var package = new ExcelPackage("path/to/file.xlsx");
var worksheet = package.Workbook.Worksheets.First(sheet => sheet.Name == "Folk");
using var reader = new CsvReader(new ExcelParser(worksheet));
var people = reader.GetRecords<Person>();
```

When an instance of `Stream` is passed to the constructor then disposal will not be handled by the parser. By default the first worksheet is used as the data source.
```csharp
var bytes = File.ReadAllBytes("path/to/file.xlsx");
using var stream = new MemoryStream(bytes);
using var parser = new ExcelParser(stream);
using var reader = new CsvReader(parser);
var people = reader.GetRecords<Person>();
// do other stuff with workbook
```

All constructor options have overloads allowing you to specify your own `Configuration`, otherwise the default is used.

## ExcelSerializer
`ExcelSerializer` implements `ISerializer` and, like `ExcelParser`, allows you to specify the path to which to (eventually) save the workbook, pass an instance of `ExcelPackage` that you have already created, or pass a specific instance of `ExcelWorksheet`, `ExcelRange` or `Stream` to use as the destination.

When the path is passed to the constructor the creation and disposal of both the workbook and worksheet (named "Export" by default) as well as the saving of the workbook on dispose, is handled by the serialiser.
```csharp
using var writer = new CsvWriter(new ExcelSerializer("path/to/file.xlsx"));
writer.WriteRecords(people);
```

When an instance of `ExcelPackage` is passed to the constructor, the creation and disposal of a new worksheet (named "Export" by default) is handled by the serialiser, but the workbook will not be saved automatically.
```csharp
using var package = new ExcelPackage();
// do stuff with the package
using var writer = new CsvWriter(new ExcelSerializer(package));
writer.WriteRecords(people);
// do other stuff with package
package.SaveAs(new FileInfo("path/to/file.xlsx"));
```

When instances of `ExcelPackage` and `ExcelWorksheet` are passed to the constructor then the serialiser will not dispose or automatically save anything.
```csharp
using var package = new ExcelPackage();
var worksheet = package.Workbook.Worksheets.Add("Folk");
using var writer = new CsvWriter(new ExcelSerializer(package, worksheet));
writer.WriteRecords(people);
package.SaveAs(new FileInfo("path/to/file.xlsx"));
```

When an instance of `Stream` is passed to the constructor the creation and disposal of a new worksheet (named "Export" by default) as well as the saving of the workbook on dispose, is handled by the serialiser. The stream however will not be disposed.
```csharp
using var stream = new MemoryStream();
using var serialiser = new ExcelSerializer(stream);
using var writer = new CsvWriter(serialiser);
writer.WriteRecords(people);
//other stuff
var bytes = stream.ToArray();
```

All constructor options have overloads allowing you to specify your own `CsvConfiguration`, otherwise the default is used.
