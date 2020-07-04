using System;
using System.Collections.Generic;
using System.Text;

namespace Flexerant.DataExport.Excel
{
    [AttributeUsage(AttributeTargets.Property)]
    public class ExcelSpreadsheetColumnAttribute : Attribute
    {
        public string ColumnName { get; set; }
        public int Order { get; set; } = int.MaxValue;

        public ExcelSpreadsheetColumnAttribute(string propertyName = null)
        {
            this.ColumnName = propertyName;
        }
    }
}
