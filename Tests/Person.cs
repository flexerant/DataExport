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
        public int Age { get; set; }

        [ExcelSpreadsheetColumn("Female", Order = 4)]
        public bool IsFemale { get; set; }

        [ExcelCellFormat(ExcelCellFormatAttribute.Accounting)]
        public double Worth { get; set; }

        [ExcelCellFormat(ExcelCellFormatAttribute.Percentage)]
        public double Percent { get; set; }

        public string Text { get; set; }

        [ExcelCellFormat(ExcelCellFormatAttribute.Text)]
        public int Integer { get; set; }
    }
}
