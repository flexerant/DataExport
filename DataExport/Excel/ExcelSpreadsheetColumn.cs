using System;
using System.Collections.Generic;
using System.Reflection;
using System.Text;

namespace Flexerant.DataExport.Excel
{
    class ExcelSpreadsheetColumn
    {
        public int Order { get; private set; } = int.MaxValue;
        public string ColumnName { get; private set; }
        public string PropertyName { get; private set; }      
        public string CellFormat { get; private set; } = null;

        public ExcelSpreadsheetColumn(PropertyInfo pi)
        {
            var colAtt = pi.GetCustomAttribute<ExcelSpreadsheetColumnAttribute>();
            var formatAtt = pi.GetCustomAttribute<ExcelCellFormatAttribute>();
            this.PropertyName = pi.Name;

            if (colAtt == null)
            {
                this.ColumnName = pi.Name;
            }
            else
            {
                if (colAtt.ColumnName == null)
                {
                    this.ColumnName = pi.Name;
                }
                else
                {
                    this.ColumnName = colAtt.ColumnName;
                }
            }

            if (formatAtt != null)
            {
                this.CellFormat = formatAtt.Format;
            }
        }
    }
}
