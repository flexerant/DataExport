using Org.BouncyCastle.Bcpg;
using System;
using System.Collections.Generic;
using System.Text;

namespace Flexerant.DataExport.Excel
{   
    public static class CellFormats
    {
        public const string SHORT_DATE = "yyyy-mm-dd";
        public const string LONG_DATE = "[$-x-sysdate]dddd, mmmm dd, yyyy";
        public const string CURRENCY = "$#,##0.00";
        public const string ACCOUNTING = "_-$* #,##0.00_-;-$* #,##0.00_-;_-$* \"-\"??_-;_-@_-";
        public const string PERCENTAGE = "0.00%";
        public const string NUMBER = "0.00";
        public const string TEXT = "@";
    }
}
