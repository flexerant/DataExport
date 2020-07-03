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
        private readonly Dictionary<string, short> _dataFormats = new Dictionary<string, short>();
        private readonly Dictionary<string, ICellStyle> _styles = new Dictionary<string, ICellStyle>();

        public ExcelWorkbook()
        {
            _workbook = new XSSFWorkbook();
        }

        public void AddSpreadheet<T>(ICollection<T> collection, string spreadsheetName = null)
        {
            Type type = typeof(T);
            PropertyInfo[] itemProperties = type.GetProperties();
            List<ExcelSpreadsheetColumn> columns = itemProperties.Select(pi => new ExcelSpreadsheetColumn(pi)).OrderBy(c => c.Order).ThenBy(c => c.ColumnName).ToList();
            ExcelSpreadsheetAttribute spreadhseetAttribute = type.GetCustomAttribute<ExcelSpreadsheetAttribute>();

            if (spreadsheetName == null)
            {
                if (spreadhseetAttribute == null)
                {
                    spreadsheetName = type.Name;
                }
                else
                {
                    spreadsheetName = spreadhseetAttribute.SpreadsheetName;
                }
            }
                                   
            var spreadsheetIndex = _workbook.GetSheetIndex(spreadsheetName);
            int sheetIndex = 0;

            while(spreadsheetIndex != -1)
            {
                sheetIndex++;
                spreadsheetName = $"{spreadsheetName}({sheetIndex})";
                spreadsheetIndex = _workbook.GetSheetIndex(spreadsheetName);
            }

            ICellStyle headerStyle = _workbook.CreateCellStyle();

            headerStyle.BorderBottom = BorderStyle.Thin;
            headerStyle.Alignment = HorizontalAlignment.Center;

            IFont headerFont = _workbook.CreateFont();

            headerFont.IsBold = true;
            headerStyle.SetFont(headerFont);
            
            int colCount = columns.Count;
            ISheet sheet = _workbook.CreateSheet(spreadsheetName);
            int rowIndex = 0;
            IRow row = sheet.CreateRow(rowIndex);

            for (int cellIndex = 0; cellIndex < colCount; cellIndex++)
            {
                ExcelSpreadsheetColumn col = columns[cellIndex];
                ICell cell = row.CreateCell(cellIndex, CellType.String);

                cell.SetCellValue(col.ColumnName);
                cell.CellStyle = headerStyle;                
            }

            rowIndex++;

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

            for (int cellIndex = 0; cellIndex < colCount; cellIndex++)
            {
                ExcelSpreadsheetColumn col = columns[cellIndex];

                sheet.AutoSizeColumn(cellIndex);
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
                Type type = value.GetType();

                switch (value)
                {
                    case int i:
                    case short s:
                    case long l:
                    case decimal dec:
                    case double dbl:
                    case float f:
                        cell.SetCellValue(Convert.ToDouble(value));
                        break;

                    case bool b:
                        cell.SetCellValue((bool)value);
                        break;

                    case DateTime dt:
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
