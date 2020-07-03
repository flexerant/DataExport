using System;
using System.Collections.Generic;
using System.Text;

namespace Flexerant.DataExport.Excel
{
    [AttributeUsage(AttributeTargets.Class)]
    public class ExcelSpreadsheetAttribute : Attribute
    {
        public string SpreadsheetName { get; private set; }

        public ExcelSpreadsheetAttribute(string spreadsheetName)
        {
            this.SpreadsheetName = spreadsheetName;
        }
    }
}
