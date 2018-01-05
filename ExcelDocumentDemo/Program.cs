using System;
using System.IO;

namespace ExcelDocumentDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            //Read XLS (Excel 97/2003)
            var path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Data", "Sample Older XLS file.xls");
            var excelModel = new ReadXLS(path, 1).Go();

            //Read XLSX (Excel 2007/2010+)
            var path2 = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Data", "Sample XLSX file.xlsx");
            var excelModel2 = new ReadXLSX(path2, 1).Go();

            //TODO - write to XLSX documents

            Console.Read();
        }
    }
}
