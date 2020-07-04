using DocumentFormat.OpenXml.Office.CustomUI;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using Org.BouncyCastle.Asn1.X509.Qualified;
using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;

namespace Flexerant.DataExport.Excel
{
    public class ExcelWorkbook
    {
        private readonly IWorkbook _workbook;
        private readonly Dictionary<string, ICellStyle> _styles = new Dictionary<string, ICellStyle>();

        public ExcelWorkbook()
        {
            _workbook = new XSSFWorkbook();
        }

      

        public void AddSpreadheet<T>(IEnumerable<T> collection, string spreadsheetName = null)
        {           
            List<ExcelSpreadsheetColumn> columns = this.GetSpreadsheetColumns<T>();           

            if (spreadsheetName == null)
            {
                spreadsheetName = this.GetSpreadsheetName<T>();
            }

            var spreadsheetIndex = _workbook.GetSheetIndex(spreadsheetName);
            int sheetIndex = 0;

            // Prevent duplicate spreadsheet names.
            while (spreadsheetIndex != -1)
            {
                sheetIndex++;
                spreadsheetName = $"{spreadsheetName}({sheetIndex})";
                spreadsheetIndex = _workbook.GetSheetIndex(spreadsheetName);
            }

            ICellStyle headerStyle = this.GetHeaderStyle();          
            int colCount = columns.Count;
            ISheet sheet = _workbook.CreateSheet(spreadsheetName);
            int rowIndex = 0;
            IRow row = sheet.CreateRow(rowIndex);

            // Set the header.
            for (int cellIndex = 0; cellIndex < colCount; cellIndex++)
            {
                ExcelSpreadsheetColumn col = columns[cellIndex];
                ICell cell = row.CreateCell(cellIndex, CellType.String);

                cell.SetCellValue(col.ColumnName);
                cell.CellStyle = headerStyle;
            }

            rowIndex++;

            // Set the remaining rows.
            foreach (var item in collection)
            {
                row = sheet.CreateRow(rowIndex);

                for (int cellIndex = 0; cellIndex < colCount; cellIndex++)
                {
                    ExcelSpreadsheetColumn col = columns[cellIndex];
                    ICell cell = row.CreateCell(cellIndex, CellType.String);
                    PropertyInfo pi = item.GetType().GetProperty(col.PropertyName);

                    this.SetCellValue(cell, pi.GetValue(item));

                    cell.CellStyle = this.GetCellStyle(col.CellFormat);
                }

                rowIndex++;
            }

            // Autosize the columns.
            for (int cellIndex = 0; cellIndex < colCount; cellIndex++)
            {
                ExcelSpreadsheetColumn col = columns[cellIndex];

                sheet.AutoSizeColumn(cellIndex);
            }
        }

        private ICellStyle GetHeaderStyle()
        {
            ICellStyle headerStyle = _workbook.CreateCellStyle();

            headerStyle.BorderBottom = BorderStyle.Thin;
            headerStyle.Alignment = HorizontalAlignment.Center;

            IFont headerFont = _workbook.CreateFont();

            headerFont.IsBold = true;
            headerStyle.SetFont(headerFont);

            return headerStyle;
        }

        private List<ExcelSpreadsheetColumn> GetSpreadsheetColumns<T>()
        {
            return typeof(T).GetProperties()
                .Select(pi => new ExcelSpreadsheetColumn(pi))
                .Where(c => !c.Ignore)
                .OrderBy(c => c.Order)
                .ThenBy(c => c.ColumnName)
                .ToList();
        }

        private string GetSpreadsheetName<T>()
        {
            Type type = typeof(T);
            ExcelSpreadsheetAttribute spreadhseetAttribute = type.GetCustomAttribute<ExcelSpreadsheetAttribute>();

            if (spreadhseetAttribute == null)
            {
                return type.Name;
            }
            else
            {
                return spreadhseetAttribute.SpreadsheetName;
            }
        }

        private ICellStyle GetCellStyle(string format)
        {
            if (string.IsNullOrWhiteSpace(format)) format = string.Empty;

            if (!_styles.ContainsKey(format))
            {
                ICellStyle style = _workbook.CreateCellStyle();

                style.Alignment = HorizontalAlignment.Center;

                if (!string.IsNullOrWhiteSpace(format))
                {
                    style.DataFormat = _workbook.CreateDataFormat().GetFormat(format);
                }

                _styles.Add(format, style);
            }

            return _styles[format];
        }

        public void Save(FileInfo fi)
        {
            using (FileStream fs = new FileStream(fi.FullName, FileMode.Create, FileAccess.Write, FileShare.Write))
            {
                this.Save(fs);
            }
        }

        public void Save(Stream stream)
        {
            _workbook.Write(stream);
        }

        private void SetCellValue(ICell cell, object value)
        {
            if (value != null)
            {
                switch (value)
                {
                    case int _:
                    case short _:
                    case long _:
                    case decimal _:
                    case double _:
                    case float _:
                        cell.SetCellValue(Convert.ToDouble(value));
                        break;

                    case bool _:
                        cell.SetCellValue((bool)value);
                        break;

                    case DateTime _:
                        cell.SetCellValue((DateTime)value);
                        break;

                    default:
                        cell.SetCellValue(value.ToString());
                        break;
                }
            }
        }
    }
}
