using System;
using System.Collections.Generic;
using System.Text;

namespace Flexerant.DataExport.Excel
{
    [AttributeUsage(AttributeTargets.Property)]
    public class ExcelCellFormatAttribute : Attribute
    {
        public const string ShortDate = "yyyy-mm-dd";
        public const string LongDate = "[$-x-sysdate]dddd, mmmm dd, yyyy";
        public const string Currency = "$#,##0.00";
        public const string Accounting = "_-$* #,##0.00_-;-$* #,##0.00_-;_-$* \"-\"??_-;_-@_-";
        public const string Percentage = "0.00%";
        public const string Number = "0.00";
        public const string Text = "@";

        public string Format { get; private set; }

        public ExcelCellFormatAttribute(string format)
        {
            this.Format = format;
        }
    }
}
