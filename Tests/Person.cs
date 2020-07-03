using Flexerant.DataExport.Excel;
using System;
using System.Collections.Generic;
using System.Security.Cryptography;
using System.Text;

namespace Tests
{
    [ExcelSpreadsheet("People")]
    public class Person
    {
        [ExcelSpreadsheetColumn("First name", Order = 0)]
        public string FirstName { get; set; }

        [ExcelSpreadsheetColumn("Last name", Order = 1)]
        public string LastName { get; set; }

        [ExcelSpreadsheetColumn("Date of birth", Order = 2)]
        [ExcelCellFormat(ExcelCellFormatAttribute.ShortDate)]
        public DateTime BirthDate { get; set; }

        [ExcelSpreadsheetColumn(Order = 3)]
        public int Age => Convert.ToInt32(Math.Floor(DateTime.Now.Subtract(this.BirthDate).TotalDays / 365));

        [ExcelSpreadsheetColumn("Female", Order = 4)]
        public bool IsFemale { get; set; }

        [ExcelSpreadsheetIgnoreColumn()]
        public Guid UUID { get; set; } = new Guid();

        [ExcelCellFormat(ExcelCellFormatAttribute.Accounting)]
        public double Worth { get; set; }

        [ExcelCellFormat(ExcelCellFormatAttribute.Percentage)]
        public double Percent { get; set; }

        public string Text { get; set; }

        [ExcelCellFormat(ExcelCellFormatAttribute.Text)]
        public int Integer { get; set; }
    }
}
