# CsvHelper for Excel (using EPPlus)

[![License](https://img.shields.io/github/license/lethek/CsvHelper.Excel.EPPlus?label=License)](https://github.com/lethek/CsvHelper.Excel.EPPlus/blob/master/LICENSE)
[![Build & Publish](https://github.com/lethek/CsvHelper.Excel.EPPlus/actions/workflows/dotnet.yml/badge.svg)](https://github.com/lethek/CsvHelper.Excel.EPPlus/actions/workflows/dotnet.yml)
[![NuGet](https://img.shields.io/nuget/v/CsvHelper.Excel.EPPlus?label=NuGet%20%28EPPlus6%29)](https://www.nuget.org/packages/CsvHelper.Excel.EPPlus)
[![NuGet](https://img.shields.io/nuget/v/CsvHelper.Excel.EPPlus?label=NuGet%20%28EPPlus4%29)](https://www.nuget.org/packages/CsvHelper.Excel.EPPlus4)

## Overview

*CsvHelper for Excel (using EPPlus)* is an extension that links two excellent libraries: [CsvHelper](https://joshclose.github.io/CsvHelper/) and [EPPlus](https://www.epplussoftware.com/).

It provides implementations of `IParser` and `IWriter` from [CsvHelper](https://joshclose.github.io/CsvHelper/) that read and write Excel documents using [EPPlus](https://www.epplussoftware.com/). Encrypted/password-protected Excel documents are supported.

&nbsp;

---

## Setup

You have a choice of two packages. It'll probably come down to your licensing requirements:
* ***[CsvHelper.Excel.EPPlus](https://www.nuget.org/packages/CsvHelper.Excel.EPPlus)*** depends on [EPPlus 6](https://github.com/EPPlusSoftware/EPPlus). This version of EPPlus has a **[Polyform Noncommercial license](https://spdx.org/licenses/PolyForm-Noncommercial-1.0.0.html)** *OR* requires you to obtain a commercial license from EPPlus Software: https://www.epplussoftware.com/LicenseOverview
* ***[CsvHelper.Excel.EPPlus4](https://www.nuget.org/packages/CsvHelper.Excel.EPPlus4)*** depends on [EPPlus 4](https://github.com/JanKallman/EPPlus). This version of EPPlus is **[LGPL](https://spdx.org/licenses/LGPL-3.0-only.html)** licensed. Consider this version if the other one is not available for your use.

Install the appropriate package from [NuGet.org](https://www.nuget.org/packages?q=CsvHelper.Excel.EPPlus) into your project. E.g.:

```
dotnet add package CsvHelper.Excel.EPPlus
```

Or using the Package Manager Console with the following command:

```
PM> Install-Package CsvHelper.Excel.EPPlus
```

Add the `CsvHelper.Excel.EPPlus` namespace to your code and check the examples below.

If you need to parse or write to a password-protected Excel document you will need to first create an instance of `ExcelPackage` yourself (e.g. `new ExcelPackage("file.xlsx", password)`) and then use one of the constructor overloads described below which take that as a parameter.

&nbsp;

---

## Using ExcelParser

`ExcelParser` implements `IParser` and allows you to specify the path of an Excel package, pass an instance of `ExcelPackage`, `ExcelWorkbook`, `ExcelWorksheet`, `ExcelRange` or a `Stream` that you have already loaded to use as the data source.

All constructor overloads have an optional parameter for passing your own `CsvConfiguration` (`IParserConfiguration`), otherwise a default constructed using the InvariantCulture is used.

&nbsp;

### **Loading records from an Excel document path**

Constructor: `ExcelParser(string path, string sheetName = null, IParserConfiguration configuration = null)`

By default the first worksheet is used as the data source, though you can specify a particular worksheet using the sheetName parameter.

When the path is passed to the constructor then workbook loading and disposal is completely handled internally by the parser.

```csharp
using var reader = new CsvReader(new ExcelParser("path/to/file.xlsx"));
var people = reader.GetRecords<Person>();
```

&nbsp;

### **Loading records from a Stream**

Constructor: `ExcelParser(Stream stream, string sheetName = null, IParserConfiguration configuration = null, bool leaveOpen = false)`

By default the first worksheet is used as the data source, though you can specify a particular worksheet using the sheetName parameter.

Unless you set `leaveOpen` to true, disposing `ExcelParser` will also automatically dispose the provided `Stream`.

```csharp
using var reader = new CsvReader(new ExcelParser(File.Open("path/to/file.xlsx", FileMode.Open)));
var people = reader.GetRecords<Person>();
```

Or explicitly managing all the dependency lifetimes rather than relying on the library to do it:

```csharp
using var stream = File.Open("path/to/file.xlsx", FileMode.Open);
using var parser = new ExcelParser(stream, leaveOpen:true);
using var reader = new CsvReader(parser, leaveOpen:true);
var people = reader.GetRecords<Person>();
```

&nbsp;

### **Loading records from an ExcelPackage**

Constructor: `ExcelParser(ExcelPackage package, string sheetName = null, IParserConfiguration configuration = null, bool leaveOpen = false)`

By default the first worksheet is used as the data source, though you can specify a particular worksheet using the sheetName parameter.

Unless you set `leaveOpen` to true, disposing `ExcelParser` will also automatically dispose the provided `ExcelPackage`.

```csharp
using var reader = new CsvReader(new ExcelParser(new ExcelPackage("path/to/file.xlsx")));
var people = reader.GetRecords<Person>();
```

Or explicitly managing all the dependency lifetimes rather than relying on the library to do it:

```csharp
using var package = new ExcelPackage("path/to/file.xlsx");
using var parser = new ExcelParser(package, leaveOpen:true);
using var reader = new CsvReader(parser, leaveOpen:true);
var people = reader.GetRecords<Person>();
```

&nbsp;

### **Loading records from an ExcelWorkbook**

Constructor: `ExcelParser(ExcelWorkbook workbook, string sheetName = null, IParserConfiguration configuration = null, bool leaveOpen = false)`

By default the first worksheet is used as the data source, though you can specify a particular worksheet using the sheetName parameter.

Unless you set `leaveOpen` to true, disposing `ExcelParser` will also automatically dispose the provided `ExcelWorkbook`.

With this overload, `ExcelParser` has no access to, or even knowledge of, the `ExcelPackage` which the `workbook` belongs to so you still need to ensure the `ExcelPackage` is appropriately disposed.

```csharp
using var package = new ExcelPackage("path/to/file.xlsx");
using var reader = new CsvReader(new ExcelParser(package.Workbook));
var people = reader.GetRecords<Person>();
```

Or explicitly managing all the dependency lifetimes rather than relying on the library to do it:

```csharp
using var package = new ExcelPackage("path/to/file.xlsx");
using var parser = new ExcelParser(package.Workbook, leaveOpen:true);
using var reader = new CsvReader(parser, leaveOpen:true);
var people = reader.GetRecords<Person>();
```

&nbsp;

### **Loading records from an ExcelWorksheet**

Constructor: `ExcelParser(ExcelWorksheet worksheet, IParserConfiguration configuration = null, bool leaveOpen = false)`

Unless you set `leaveOpen` to true, disposing `ExcelParser` will also automatically dispose the `ExcelWorkbook` that owns the provided `ExcelWorksheet`.

With this overload, `ExcelParser` has no access to, or even knowledge of, the `ExcelPackage` which the `worksheet` belongs to so you still need to ensure the `ExcelPackage` is appropriately disposed.

```csharp
using var package = new ExcelPackage("path/to/file.xlsx");
using var reader = new CsvReader(new ExcelParser(package.Workbook.Worksheets.First(sheet => sheet.Name == "Folk")));
var people = reader.GetRecords<Person>();
```

Or explicitly managing all the dependency lifetimes rather than relying on the library to do it:

```csharp
using var package = new ExcelPackage("path/to/file.xlsx");
var worksheet = package.Workbook.Worksheets.First(sheet => sheet.Name == "Folk");
using var parser = new ExcelParser(worksheet, leaveOpen:true);
using var reader = new CsvReader(parser, leaveOpen:true);
var people = reader.GetRecords<Person>();
```

&nbsp;

### **Loading records from an ExcelRange**

Constructor: `ExcelParser(ExcelRange range, IParserConfiguration configuration = null, bool leaveOpen = false)`

This overload allows you to restrict the parsing to a specific range of cells within an Excel worksheet.

With this overload, `ExcelParser` has no access to, or even knowledge of, the `ExcelPackage` which the `range` belongs to so you still need to ensure the `ExcelPackage` is appropriately disposed.

Unless you set `leaveOpen` to true, disposing `ExcelParser` will also automatically dispose the `ExcelWorkbook` that owns the provided `ExcelRange`.

```csharp
using var package = new ExcelPackage("path/to/file.xlsx");
var range = package.Workbook.Worksheets.First(sheet => sheet.Name == "Folk").Cells[2, 5, 400, 33];
using var reader = new CsvReader(new ExcelParser(range));
var people = reader.GetRecords<Person>();
```

Or explicitly managing all the dependency lifetimes rather than relying on the library to do it:

```csharp
using var package = new ExcelPackage("path/to/file.xlsx");
var range = package.Workbook.Worksheets.First(sheet => sheet.Name == "Folk");
using var parser = new ExcelParser(range, leaveOpen:true);
using var reader = new CsvReader(parser, leaveOpen:true);
var people = reader.GetRecords<Person>();
```

&nbsp;

---

## Using ExcelWriter

`ExcelWriter` implements `IWriter` and, like `ExcelParser`, allows you to specify the path to (eventually) save the workbook, pass an instance of `ExcelPackage` that you have already created, or pass a specific instance of `ExcelWorksheet`, `ExcelRange` or `Stream` to use as the destination.

Unlike `ExcelParser` and `CsvReader` however where CsvReader wraps ExcelParser, here `ExcelWriter` inherits from `CsvWriter` and should be used directly instead.

All constructor overloads have an optional parameter for passing your own `CsvConfiguration` (`IWriterConfiguration`), otherwise a default constructed using the InvariantCulture is used.

&nbsp;

### **Writing records to an Excel document path**

Constructor: `ExcelWriter(string path, string sheetName = "Export", IWriterConfiguration configuration = null)`

When the path is passed to the constructor the writer manages the creation & disposal of the workbook and worksheet (named "Export" by default). The workbook is saved only when the writer is disposed.

```csharp
using var writer = new ExcelWriter("path/to/file.xlsx");
writer.WriteRecords(people);
```

&nbsp;

### **Writing records to a Stream**

Constructor: `ExcelWriter(Stream stream, string sheetName = "Export", IWriterConfiguration configuration = null, bool leaveOpen = false)`

Important: The data is saved only when the `ExcelWriter` is disposing.

Unless you set `leaveOpen` to true, disposing `ExcelWriter` will also automatically dispose the provided `Stream`.

```csharp
using var writer = new ExcelWriter(new MemoryStream());
writer.WriteRecords(people);
```

&nbsp;

### **Writing records to an ExcelPackage**

Constructor: `ExcelWriter(ExcelPackage package, string sheetName = "Export", IWriterConfiguration configuration = null, bool leaveOpen = false)`

Important: The data is saved only when the `ExcelWriter` is disposing or the consumer manually calls `package.Save()` or `package.SaveAs(...)`.

By default, records are written into a worksheet named "Export".

Unless you set `leaveOpen` to true, disposing `ExcelWriter` will also automatically dispose the provided `ExcelPackage`.

```csharp
using var writer = new ExcelWriter(new ExcelPackage());
writer.WriteRecords(people);
package.SaveAs("path/to/file.xlsx");
```

Or

```csharp
using var writer = new ExcelWriter(new ExcelPackage("path/to/file.xlsx"));
writer.WriteRecords(people);
```

&nbsp;

### **Writing records to an ExcelWorksheet**

Constructor: `ExcelWriter(ExcelPackage package, ExcelWorksheet worksheet, IWriterConfiguration configuration = null, bool leaveOpen = false)`

Important: The data is saved only when the `ExcelWriter` is disposing or the consumer manually calls `package.Save()` or `package.SaveAs(...)`.

This overload is the same as the one which takes `ExcelPackage` and `sheetName` parameters, but accepts a worksheet reference rather than name.

Unless you set `leaveOpen` to true, disposing `ExcelWriter` will also automatically dispose the provided `ExcelPackage`.

```csharp
using var package = new ExcelPackage();
var worksheet = package.Workbook.Worksheets.Add("Folk");
using var writer = new ExcelWriter(package, worksheet);
writer.WriteRecords(people);
package.SaveAs("path/to/file.xlsx");
```

Or

```csharp
using var package = new ExcelPackage("path/to/file.xlsx");
var worksheet = package.Workbook.Worksheets.Add("Folk");
using var writer = new ExcelWriter(package, worksheet);
writer.WriteRecords(people);
```

&nbsp;

### **Writing records to an ExcelRange**

Constructor: `ExcelWriter(ExcelPackage package, ExcelRange range, IWriterConfiguration configuration = null, bool leaveOpen = false)`

Important: The data is saved only when the `ExcelWriter` is disposing or the consumer manually calls `package.Save()` or `package.SaveAs(...)`.

This overload is similar to the previous ones but accepts an `ExcelRange` instead, allowing targeting a specific range of cells within an Excel worksheet.

Unless you set `leaveOpen` to true, disposing `ExcelWriter` will also automatically dispose the provided `ExcelPackage`.

```csharp
using var package = new ExcelPackage();
var worksheet = package.Workbook.Worksheets.Add("Folk");
using var writer = new ExcelWriter(package, worksheet.Cells[2, 5, 400, 33]);
writer.WriteRecords(people);
package.SaveAs("path/to/file.xlsx");
```

Or

```csharp
using var package = new ExcelPackage("path/to/file.xlsx");
var worksheet = package.Workbook.Worksheets.Add("Folk");
using var writer = new ExcelWriter(package, worksheet.Cells[2, 5, 400, 33]);
writer.WriteRecords(people);
```

&nbsp;

---

## Attribution

***This project was originally forked from https://github.com/christophano/CsvHelper.Excel and https://github.com/youngcm2/CsvHelper.Excel and heavily modified so that it could be used with [EPPlus](https://www.nuget.org/packages/EPPlus) instead of ClosedXml.***
