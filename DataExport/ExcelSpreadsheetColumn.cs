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
        public bool Ignore { get; private set; } = false;

        public ExcelSpreadsheetColumn(PropertyInfo pi)
        {
            var colAtt = pi.GetCustomAttribute<ExcelSpreadsheetColumnAttribute>();
            var ignoreAtt = pi.GetCustomAttribute<ExcelSpreadsheetIgnoreColumnAttribute>();

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

                this.Order = colAtt.Order;
                this.CellFormat = colAtt.CellFormat;
            }

            if (ignoreAtt != null)
            {
                this.Ignore = true;
            }
        }
    }
}
