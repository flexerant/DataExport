using System;

namespace Flexerant.DataExport
{
    public class ExcelColumnAttribute : Attribute
    {
        public string ColumnName { get; private set; }

        public ExcelColumnAttribute(string columnName)
        {
            this.ColumnName = columnName.ToUpper();
        }

        public int GetColumnNumber()
        {
            return GetColumnNumber(this.ColumnName);
        }

        //* Kudos: https://stackoverflow.com/a/848184

        public static int GetColumnNumber(string columnName)
        {
            int number = 0;
            int pow = 1;

            for (int i = columnName.Length - 1; i >= 0; i--)
            {
                number += (columnName[i] - 'A' + 1) * pow;
                pow *= 26;
            }

            return number - 1;
        }
    }
}
