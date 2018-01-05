namespace ExcelDocumentDemo.Common
{
    public class FormatCellNumber
    {
        public static string ToInternational(object value, int prefix = 27)
        {
            var v = value.ToString().Replace(" ", "");
            if (v.StartsWith("0"))
                return string.Format("27{0}", v.Substring(1, v.Length - 1));            
            else
                return v;
        }
    }
}
