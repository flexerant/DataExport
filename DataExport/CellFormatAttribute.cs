using System;
using System.Collections.Generic;
using System.Text;

namespace Flexerant.DataExport
{
    [AttributeUsage(AttributeTargets.Property | AttributeTargets.Field)]
    public class CellFormatAttribute : Attribute
    {        
        public ReportCellFormats CellFormat { get; private set; }
       
        public CellFormatAttribute(ReportCellFormats cellFormat)
        {
            this.CellFormat = cellFormat;
        }
    }
}
