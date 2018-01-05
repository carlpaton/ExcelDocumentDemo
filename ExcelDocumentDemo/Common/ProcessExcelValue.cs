using System;

namespace ExcelDocumentDemo.Common
{
    public class ProcessExcelValue
    {
        public static void Go(int cellCount, object val, ExcelDataModel excelModel)
        {
            switch (cellCount) //Excel columns should be known
            {
                case 1:
                    excelModel.Id = Convert.ToInt32(val);
                    break;
                case 2:
                    excelModel.Surname = val.ToString();
                    break;
                case 3:
                    excelModel.FirstName = val.ToString();
                    break;
                case 4:
                    excelModel.Cell = Convert.ToInt64(FormatCellNumber.ToInternational(val));
                    break;
                case 5:
                    excelModel.Email = val.ToString();
                    break;
                case 6:
                    if (val != null)
                    {
                        if (val.ToString() != "")
                            excelModel.StartDate = DateTime.ParseExact(val.ToString(), "dd/MM/yyyy", System.Globalization.CultureInfo.CurrentCulture);
                    }
                    break;
            }
        }
    }
}
