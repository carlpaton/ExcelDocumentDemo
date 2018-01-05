using System.Collections.Generic;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System;
using ExcelDocumentDemo.Common;
using System.IO;

namespace ExcelDocumentDemo
{
    /// <summary>
    /// .xls are older excel documents, these should be migrated over to the newer .xlsx format
    /// This class is for legacy offce 97/2003 support AND OR dumb users :D
    /// </summary>
    public class ReadXLS
    {
        private string ExcelPath = "";
        private int WorkSheetToRead = 1;
        private bool IgnoreHeader = true; //format is known, else add to constructor

        public ReadXLS(string excelPath, int workSheetToRead)
        {
            ExcelPath = excelPath;
            WorkSheetToRead = workSheetToRead;
        }

        public List<ExcelDataModel> Go()
        {
            var r = new List<ExcelDataModel>();

            if (ExcelPath == "")
                return r;

            if (WorkSheetToRead < 1)
                return r;

            if (!File.Exists(ExcelPath))
                return r;

            var app = new Application();

            var wb = app.Workbooks.Open(
                ExcelPath, 0, true, 5, "", "", true, 
                XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);

            var ws = (Worksheet)wb.Worksheets.get_Item(WorkSheetToRead);
            var range = ws.UsedRange;
            int rowCount, cellCount;

            for (rowCount = 1; rowCount <= range.Rows.Count; rowCount++)
            {
                if (IgnoreHeader)
                {
                    IgnoreHeader = false;
                    continue;
                }

                var excelModel = new ExcelDataModel();
                for (cellCount = 1; cellCount <= range.Columns.Count; cellCount++)
                {
                    try
                    {
                        var v = (range.Cells[rowCount, cellCount] as Range).Value2;
                        if (cellCount == 6)
                        {
                            //format date from 99999 to dd/MM/yyyy
                            var x = DateTime.FromOADate(v).ToShortDateString();
                            ProcessExcelValue.Go(cellCount, x, excelModel);
                        }
                        else 
                            ProcessExcelValue.Go(cellCount, v, excelModel);
                    }
                    catch (Exception ex)
                    {
                        //TODO - log
                        Console.WriteLine(ex.Message);
                    }
                }
                r.Add(excelModel);
            }

            wb.Close(false, null, null); //no need to save the document as we just reading it
            app.Quit();

            Marshal.ReleaseComObject(ws);
            Marshal.ReleaseComObject(wb);
            Marshal.ReleaseComObject(app);

            return r;
        }
    }
}
