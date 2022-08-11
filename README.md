[![Build Status](https://github.com/lethek/CsvHelper.Excel.EPPlus4/actions/workflows/dotnet.yml/badge.svg)](https://github.com/lethek/CsvHelper.Excel.EPPlus4/actions/workflows/dotnet.yml)

# CSV Helper for Excel

***This project has been forked from https://github.com/christophano/CsvHelper.Excel and https://github.com/youngcm2/CsvHelper.Excel and heavily modified; primarily so that it can be used with the final LGPL version of [EPPlus](https://github.com/JanKallman/EPPlus) instead of ClosedXml, because it works with encrypted/password-protected Excel documents.***

***NuGet packages of this fork are available from MyGet:  https://www.myget.org/feed/lethek/package/nuget/CsvHelper.Excel.EPPlus4***

CsvHelper for Excel is an extension that links two excellent libraries, [CsvHelper](https://joshclose.github.io/CsvHelper/) and [EPPlus](https://github.com/JanKallman/EPPlus).
It provides implementations of `IParser` and `IWriter` from [CsvHelper](https://joshclose.github.io/CsvHelper/) that read and write to Excel using [EPPlus](https://github.com/JanKallman/EPPlus).

If you need to parse or write to a password-protected Excel document you will need to create an instance of `ExcelPackage` yourself (e.g. `new ExcelPackage("file.xlsx", password)`) and use one of the constructor overloads described below which take that as a parameter.

---

## ExcelParser
`ExcelParser` implements `IParser` and allows you to specify the path of the Excel package, pass an instance of `ExcelPackage`, `ExcelWorkbook`, `ExcelWorksheet`, `ExcelRange` or a `Stream` that you have already loaded to use as the data source.

All constructor overloads have an optional parameter allowing you to specify your own `CsvConfiguration`, otherwise the default is used.

Explaining each of the constructors:

### `new ExcelParser(string path, string sheetName = null, CsvConfiguration configuration = null)`

When the path is passed to the constructor then the workbook loading and disposal is handled by the parser. By default the first worksheet is used as the data source, though you can specify a particular worksheet using the sheetName parameter.

```csharp
using var reader = new CsvReader(new ExcelParser("path/to/file.xlsx"));
var people = reader.GetRecords<Person>();
```

### `new ExcelParser(Stream stream, string sheetName = null, CsvConfiguration configuration = null)`

When an instance of `Stream` is passed to the constructor then disposal will not be handled by the parser unless an instance of `CsvConfiguration` with its `LeaveOpen` property set to `false` is also passed. By default the first worksheet is used as the data source, though you can specify a particular worksheet using the sheetName parameter.

```csharp
var bytes = File.ReadAllBytes("path/to/file.xlsx");
using var stream = new MemoryStream(bytes);
using var parser = new ExcelParser(stream);
using var reader = new CsvReader(parser);
var people = reader.GetRecords<Person>();
// do other stuff with workbook
```

### `new ExcelParser(ExcelPackage package, string sheetName = null, CsvConfiguration configuration = null)`

When an instance of `ExcelPackage` is passed to the constructor then disposal will not be handled by the parser unless an instance of `CsvConfiguration` with its `LeaveOpen` property set to `false` is also passed. By default the first worksheet is used as the data source, though you can specify a particular worksheet using the sheetName parameter.

```csharp
using var package = new ExcelPackage("path/to/file.xlsx");
// do stuff with the package
using var reader = new CsvReader(new ExcelParser(package));
var people = reader.GetRecords<Person>();
// do other stuff with workbook
```

### `new ExcelParser(ExcelWorkbook workbook, string sheetName = null, CsvConfiguration configuration = null)`
When an instance of `ExcelWorkbook` is passed to the constructor then disposal will not be handled by the parser unless an instance of `CsvConfiguration` with its `LeaveOpen` property set to `false` is also passed. By default the first worksheet is used as the data source, though you can specify a particular worksheet using the sheetName parameter.

```csharp
using var package = new ExcelPackage("path/to/file.xlsx");
// do stuff with the package
using var reader = new CsvReader(new ExcelParser(package.Workbook));
var people = reader.GetRecords<Person>();
// do other stuff with workbook
```


### `new ExcelParser(ExcelWorksheet worksheet, CsvConfiguration configuration = null)`

When an instance of `ExcelWorksheet` is passed to the constructor then disposal will not be handled by the parser and the worksheet will be used as the data source.

```csharp
using var package = new ExcelPackage("path/to/file.xlsx");
var worksheet = package.Workbook.Worksheets.First(sheet => sheet.Name == "Folk");
using var reader = new CsvReader(new ExcelParser(worksheet));
var people = reader.GetRecords<Person>();
```

### `new ExcelParser(ExcelRange range, CsvConfiguration configuration = null)`
When an instance of `ExcelRange` is passed to the constructor then disposal will not be handled by the parser and the range will be used as the data source. This overload allows you to restrict the parsing to a specific range of cells within an Excel worksheet.

```csharp
using var package = new ExcelPackage("path/to/file.xlsx");
var worksheet = package.Workbook.Worksheets.First(sheet => sheet.Name == "Folk");
using var reader = new CsvReader(new ExcelParser(worksheet.Cells[2, 5, 400, 33]));
var people = reader.GetRecords<Person>();
```

---

## ExcelWriter
`ExcelWriter` implements `IWriter` and, like `ExcelParser`, allows you to specify the path to (eventually) save the workbook, pass an instance of `ExcelPackage` that you have already created, or pass a specific instance of `ExcelWorksheet`, `ExcelRange` or `Stream` to use as the destination.

All constructor options have overloads allowing you to specify your own `CsvConfiguration`, otherwise the default is used.

### `new ExcelWriter(string path, string sheetName = "Export", CsvConfiguration configuration = null)`

When the path is passed to the constructor the writer manages the creation & disposal of the workbook and worksheet (named "Export" by default). The workbook is saved only when the writer is disposed.

```csharp
using var writer = new CsvWriter(new ExcelWriter("path/to/file.xlsx"));
writer.WriteRecords(people);
```

### `new ExcelWriter(Stream stream, string sheetName = "Export", CsvConfiguration configuration = null)`

When an instance of `Stream` is passed to the constructor the writer manages the creation & disposal of the workbook and worksheet (named "Export" by default). The workbook is saved only when the writer is disposed. As the stream is an external dependency however, it will not be automatically disposed by the writer's disposal unless an instance of `CsvConfiguration` with its `LeaveOpen` property set to `false` is also passed.

```csharp
using var stream = new MemoryStream();
using var serialiser = new ExcelWriter(stream);
using var writer = new CsvWriter(serialiser);
writer.WriteRecords(people);
//other stuff
var bytes = stream.ToArray();
```

### `new ExcelWriter(ExcelPackage package, string sheetName = "Export", CsvConfiguration configuration = null)`

When an instance of `ExcelPackage` is passed to the constructor, it will not be automatically disposed by the writer's disposal unless an instance of `CsvConfiguration` with its `LeaveOpen` property set to `false` is also passed. The workbook is saved only when the writer is disposed or the consumer manually calls `package.Save()` or `package.SaveAs(...)`.

By default, records are written into a worksheet named "Export".

```csharp
using var package = new ExcelPackage();
// do stuff with the package
using var writer = new CsvWriter(new ExcelWriter(package));
writer.WriteRecords(people);
// do other stuff with package
package.SaveAs(new FileInfo("path/to/file.xlsx"));
```

### `new ExcelWriter(ExcelPackage package, ExcelWorksheet worksheet, CsvConfiguration configuration = null)`

The same as the overload which takes `ExcelPackage` and `sheetName` parameters, but this one allows specifying the worksheet by reference rather than name. As before, the workbook is saved only when the writer is disposed or the consumer manually calls `package.Save()` or `package.SaveAs(...)`.

When the writer is disposed it will not automatically dispose the package unless an instance of `CsvConfiguration` with its `LeaveOpen` property set to `false` was also passed.

```csharp
using var package = new ExcelPackage();
var worksheet = package.Workbook.Worksheets.Add("Folk");
using var writer = new CsvWriter(new ExcelWriter(package, worksheet));
writer.WriteRecords(people);
package.SaveAs(new FileInfo("path/to/file.xlsx"));
```

### `new ExcelWriter(ExcelPackage package, ExcelRange range, CsvConfiguration configuration = null)`

Similar to the overload which takes `ExcelPackage` and `ExcelWorksheet` parameters, but this one allows targeting a specific range of cells within an Excel worksheet. As before, the workbook is saved only when the writer is disposed or the consumer manually calls `package.Save()` or `package.SaveAs(...)`.

```csharp
using var package = new ExcelPackage();
var worksheet = package.Workbook.Worksheets.Add("Folk");
using var writer = new CsvWriter(new ExcelWriter(package, worksheet.Cells[2, 5, 400, 33]));
writer.WriteRecords(people);
package.SaveAs(new FileInfo("path/to/file.xlsx"));
```
