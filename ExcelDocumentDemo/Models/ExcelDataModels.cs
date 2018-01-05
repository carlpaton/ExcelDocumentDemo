using System;

namespace ExcelDocumentDemo
{
    public class ExcelDataModel
    {
        public int Id { get; set; }
        public string FirstName { get; set; }
        public string Surname { get; set; }
        public long? Cell { get; set; }
        public string Email { get; set; }
        public DateTime? StartDate { get; set; } 
    }
}
