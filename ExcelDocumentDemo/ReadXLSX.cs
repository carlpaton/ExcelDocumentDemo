using ExcelDocumentDemo.Common;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace ExcelDocumentDemo
{
    /// <summary>
    /// Read XLSX files
    /// </summary>
    public class ReadXLSX
    {
        private string ExcelPath = "";
        private int WorkSheetToRead = 1;
        private bool IgnoreHeader = true; //format is known, else add to constructor

        public ReadXLSX(string excelPath, int workSheetToRead)
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

            using (var pck = new OfficeOpenXml.ExcelPackage())
            {
                using (var stream = File.OpenRead(ExcelPath))
                {
                    pck.Load(stream);
                }
                var ws = pck.Workbook.Worksheets.First(); //read the first workbook

                int totalRows = ws.Dimension.End.Row;
                int totalCols = ws.Dimension.End.Column;

                int rowStart = 1;
                if (IgnoreHeader)
                {
                    rowStart = 2;
                    IgnoreHeader = false;
                }

                try
                {
                    for (int i = rowStart; i <= totalRows; i++)
                    {
                        var excelModel = new ExcelDataModel();
                        for (int j = 1; j <= totalCols; j++)
                        {
                            var v = "";
                            if (ws.Cells[i, j].Value != null)
                                v = ws.Cells[i, j].Value.ToString();

                            if (j == 6)
                            {
                                if (v.Length >= 10)
                                    v = v.Substring(0, 10);

                                ProcessExcelValue.Go(j, v, excelModel);
                            }
                            else
                                ProcessExcelValue.Go(j, v, excelModel);

                            switch (j) //Excel columns should be known
                            {
                                case 1:
                                    excelModel.Id = Convert.ToInt32(v);
                                    break;
                                case 2:
                                    excelModel.Surname = v;
                                    break;
                                case 3:
                                    excelModel.FirstName = v;
                                    break;
                                case 4:
                                    excelModel.Cell = Convert.ToInt64(FormatCellNumber.ToInternational(v));
                                    break;
                                case 5:
                                    excelModel.Email = v;
                                    break;
                            }
                        }
                        r.Add(excelModel);
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
            }

            return r;
        }
    }
}
