# Change Log

Major version numbers below are kept in-sync with the dependent version of CsvHelper. Minor & patch version numbers below may differ from CsvHelper's.

The listed Breaking Changes are those added by this library, above-and-beyond those introduced by new versions of the CsvHelper library.

---

## 30.0.0

### Breaking Changes

* Depends on CsvHelper 30.0.1.
* Added an optional bool `leaveOpen` parameter onto all `ExcelWriter` and `ExcelParser` constructors except the two which take a path.
* By default when `ExcelWriter` and `ExcelParser` are disposed, they now also dispose of any Stream, ExcelPackage or ExcelWorkbook provided to their constructors. To keep any of these objects open, you must explicitly set the optional `leaveOpen` constructor parameter to `true`. E.g.

  ```csharp
  using var writer = new ExcelWriter(stream, leaveOpen: true);
  ```

  This brings behaviour into line with CsvHelper's defaults.

### Features

* Pass `IWriterConfiguration` into `ExcelWriter` constructor instead of `CsvConfiguration`.
* Pass `IParserConfiguration` into `ExcelParser` constructor instead of `CsvConfiguration`.

### Bug Fixes

* Fixed `ExcelParser` not disposing `ExcelPackage` when it was documented to.

---

## 29.0.0

### Breaking Changes

* Depends on CsvHelper 29.0.0.

---

## 28.1.0

### Bug Fixes

* `ExcelParser` now respects the `CsvConfiguration.TrimOptions`.
* `ExcelParser.Delimiter` now returns the configured delimiter from CsvConfiguration. This is merely for consistency with CsvParser's behaviour as the concept of a delimiter doesn't make sense for an Excel document anyway.

---

## 28.0.0

### Breaking Changes

* Depends on CsvHelper 28.0.1.

### Features

* Added support two different versions of EPPlus: `CsvHelper.Excel.EPPlus4` requires EPPlus v4, and `CsvHelper.Excel.EPPlus` requires at least EPPlus v6. Select the version based on your licensing situation.
* Upgraded unit tests to run in .NET 6 and .NET Core 3.1.
* Use GitHub Actions for CI instead of Azure Pipelines.
