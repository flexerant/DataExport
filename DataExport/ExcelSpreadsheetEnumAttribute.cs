using System;
using System.Collections.Generic;
using System.Text;

namespace Flexerant.DataExport.Excel
{
    [AttributeUsage(AttributeTargets.Field)]
    public class ExcelSpreadsheetEnumAttribute : Attribute
    {
        public object Value { get; set; } = null;

        public ExcelSpreadsheetEnumAttribute(object value = null)
        {
            this.Value = value;
        }
    }
}
