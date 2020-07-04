using System;
using System.Collections.Generic;
using System.Text;

namespace Flexerant.DataExport.Excel
{
    [AttributeUsage(AttributeTargets.Property)]
    public  class ExcelSpreadsheetIgnoreColumnAttribute : Attribute
    {
    }
}
