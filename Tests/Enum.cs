using Flexerant.DataExport.Excel;
using System;
using System.Collections.Generic;
using System.Text;

namespace Tests
{
    public enum Statuses
    {
        [ExcelSpreadsheetEnum("Pending")]
        Pending = 0,

        [ExcelSpreadsheetEnum("In progress")]
        InProgress = 1,

        [ExcelSpreadsheetEnum("Shipped")]
        Shipped = 2
    }
}
