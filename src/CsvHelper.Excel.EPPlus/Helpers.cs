using System.IO;
using System.Runtime.CompilerServices;

using OfficeOpenXml;


[assembly: InternalsVisibleTo("CsvHelper.Excel.EPPlus.Tests")]

namespace CsvHelper.Excel.EPPlus
{
    internal static class Helpers
    {
        public static ExcelPackage GetOrCreatePackage(string path, string worksheetName) {
            var file = new FileInfo(path);
            if (!file.Exists) {
                using var package = new ExcelPackage(file);
                package.GetOrAddWorksheet(worksheetName);
                package.Save();
            }
            return new ExcelPackage(file);
        }


        public static ExcelWorksheet GetOrAddWorksheet(this ExcelPackage package, string sheetName)
            => package.Workbook.Worksheets[sheetName] ?? package.Workbook.Worksheets.Add(sheetName);


        public static ExcelWorksheet GetOrAddWorksheet(this ExcelWorkbook workbook, string sheetName)
            => workbook.Worksheets[sheetName] ?? workbook.Worksheets.Add(sheetName);


        public static void Delete(string path) {
            try {
                var directory = Path.GetDirectoryName(path);
                if (Directory.Exists(directory)) {
                    Directory.Delete(directory, true);
                }
            } catch {
                //Ignore errors
            }
        }
    }
}
