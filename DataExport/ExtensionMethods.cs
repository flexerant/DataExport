using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using System.Runtime.CompilerServices;

namespace Flexerant.DataExport
{
    public static partial class ExtensionMethods
    {

        public static string GetReportPropertyName<T>(this T obj) where T : Enum
        {
            var enumType = typeof(T);
            var memberInfos = enumType.GetMember(obj.ToString());
            var enumValueMemberInfo = memberInfos.FirstOrDefault(m => m.DeclaringType == enumType);
            var valueAttributes = enumValueMemberInfo.GetCustomAttributes(typeof(ExcelSpreadsheetColumnAttribute), false);

            if (valueAttributes == null) return null;

            return ((ExcelSpreadsheetColumnAttribute)valueAttributes[0]).ColumnName;
        }

        public static string GetReportPropertyName<T, P>(this T obj, Expression<Func<T, P>> action) where T : class
        {
            var expression = (MemberExpression)action.Body;
            string name = expression.Member.Name;
            PropertyInfo pi = obj.GetType().GetProperty(name);

            if (pi == null) return null;

            var reportColumnNameProp = pi.GetCustomAttribute<ExcelSpreadsheetColumnAttribute>();

            if (reportColumnNameProp == null) return null;

            return reportColumnNameProp.ColumnName;
        }

        //public static ReportCellFormats GetCellFormat<T, P>(this T obj, Expression<Func<T, P>> action) where T : class
        //{
        //    var expression = (MemberExpression)action.Body;
        //    string name = expression.Member.Name;
        //    PropertyInfo pi = obj.GetType().GetProperty(name);

        //    if (pi == null) return ReportCellFormats.Text;

        //    var att = pi.GetCustomAttribute<CellFormatAttribute>();

        //    if (att == null) return ReportCellFormats.Text;

        //    return att.CellFormat;
        //}

        private static UInt32Value GetStyleIndex(ReportCellFormats format)
        {
            switch (format)
            {
                case ReportCellFormats.Currency:
                    return new UInt32Value(3U);

                case ReportCellFormats.Date:
                    return new UInt32Value(5U);

                case ReportCellFormats.DateTime:
                    return new UInt32Value(6U);

                case ReportCellFormats.Number:
                    return new UInt32Value(2U);

                case ReportCellFormats.Percentage:
                    return new UInt32Value(4U);

                //case ReportCellFormats.Text:
                //    return new UInt32Value(0U);

                default:
                    return new UInt32Value(1U);
            }
        }


        public static CellValues ToCellValue(this Type type)
        {
            if (type == typeof(int)) return CellValues.Number;
            if (type == typeof(short)) return CellValues.Number;
            if (type == typeof(long)) return CellValues.Number;
            if (type == typeof(decimal)) return CellValues.Number;
            if (type == typeof(double)) return CellValues.Number;
            if (type == typeof(float)) return CellValues.Number;
            if (type == typeof(bool)) return CellValues.String;
            if (type == typeof(DateTime))
            {
                return CellValues.Number;
            }

            return CellValues.String;
        }

        private static Cell ToCell(this object value)
        {
            var type = value.GetType();
            var cellValue = type.ToCellValue();

            Cell cell = new Cell();
            cell.DataType = cellValue;

            if (type == typeof(DateTime))
            {
                double oaValue = ((DateTime)value).ToOADate();


                cell.CellValue = new CellValue(oaValue.ToString(CultureInfo.InvariantCulture));
                cell.StyleIndex = GetStyleIndex(ReportCellFormats.Date);
            }
            else if (cellValue == CellValues.Number)
            {
                cell.CellValue = new CellValue(value.ToString());
                cell.StyleIndex = GetStyleIndex(ReportCellFormats.Currency);
            }
            else
            {
                cell.CellValue = new CellValue(value.ToString());
            }

            //    if (cell.DataType == CellValues.Date)
            //{
            //    double oaValue = ((DateTime)value).ToOADate();

            //    cell.CellValue = new CellValue(((DateTime)value).ToString("s"));                
            //}
            //else
            //{
            //    cell.CellValue = new CellValue(value.ToString());
            //}

            return cell;
        }

        public static DataTable ToDataTable<T>(this IEnumerable<T> values, string tableName)
        {
            DataTable table = new DataTable(tableName);

            foreach (var prop in values.First().ToProps())
            {
                Type propType = prop.PropertyType;
                Type underlyingType = Nullable.GetUnderlyingType(prop.PropertyType);

                if (underlyingType != null)
                {
                    propType = underlyingType;
                }

                table.Columns.Add(prop.PropertyName, propType);
            }

            foreach (var value in values)
            {
                var row = table.NewRow();

                foreach (var prop in value.ToProps())
                {
                    row[prop.PropertyName] = prop.PropertyValue;
                }

                table.Rows.Add(row);
            }

            return table;
        }

        private static IEnumerable<ReportProperty> ToProps<T>(this T value)
        {
            List<Tuple<int, string, Type, object, bool>> props = new List<Tuple<int, string, Type, object, bool>>();

            foreach (var pi in typeof(T).GetProperties())
            {
                var reportColumnNameProp = pi.GetCustomAttribute<ExcelSpreadsheetColumnAttribute>();

                if (reportColumnNameProp == null) continue;

                string columnName = pi.Name;
                int order = reportColumnNameProp.Order;
                Type type = pi.PropertyType;
                object piVal = pi.GetValue(value);
                Type underlyingType = Nullable.GetUnderlyingType(type);
                bool isNullable = false;

                if (underlyingType != null)
                {
                    type = underlyingType;
                    isNullable = true;
                }

                if (reportColumnNameProp.ColumnName != null) columnName = reportColumnNameProp.ColumnName;

                if (piVal == null)
                {
                    piVal = DBNull.Value;
                }

                switch (piVal)
                {
                    case short shortValue:
                        type = typeof(int);
                        piVal = Convert.ToInt32(piVal);
                        break;

                    case decimal decimalValue:
                    case float floateValue:
                        type = typeof(double);
                        piVal = Convert.ToDouble(piVal);
                        break;

                    case Enum enumValue:
                        type = typeof(string);
                        piVal = piVal.ToString();
                        break;
                }

                props.Add(new Tuple<int, string, Type, object, bool>(order, columnName, type, piVal, isNullable));
            }

            return props.OrderBy(x => x.Item1).ThenBy(x => x.Item2).Select(x => new ReportProperty() { PropertyName = x.Item2, PropertyType = x.Item3, PropertyValue = x.Item4, IsNullable = x.Item5 });
        }

        public static byte[] ToXlsx(this DataSet ds)
        {
            byte[] output = null;

            using (MemoryStream ms = new MemoryStream())
            {
                ds.ToXlsx(ms);
                output = ms.ToArray();
            }

            return output;
        }

        public static void ToXlsx(this DataSet ds, Stream stream, bool columnNamesAsHeaders = true)
        {
            using (var workbook = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook))
            {
                var workbookPart = workbook.AddWorkbookPart();

                workbook.WorkbookPart.Workbook = new Workbook();

                workbook.WorkbookPart.Workbook.Sheets = new Sheets();

                var stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
                stylesPart.Stylesheet = new Stylesheet
                {
                    Fonts = new Fonts(new Font()),
                    Fills = new Fills(new Fill()),
                    Borders = new Borders(new Border()),
                    CellStyleFormats = new CellStyleFormats(new CellFormat()),

                    // https://stackoverflow.com/a/4655716
                    // http://www.ecma-international.org/news/TC45_current_work/Office%20Open%20XML%20Part%204%20-%20Markup%20Language%20Reference.pdf
                    CellFormats =
                            new CellFormats(
                            new CellFormat(),
                             new CellFormat // general (1)
                             {
                                 NumberFormatId = 0,
                                 ApplyNumberFormat = true
                             },
                             new CellFormat // number (2)
                             {
                                 NumberFormatId = 2,
                                 ApplyNumberFormat = true
                             },
                             new CellFormat // currency (3)
                             {
                                 NumberFormatId = 7,
                                 ApplyNumberFormat = true
                             },
                            new CellFormat // percentage (4)
                            {
                                NumberFormatId = 10,
                                ApplyNumberFormat = true
                            },
                            new CellFormat // date (5)
                            {
                                NumberFormatId = 14,
                                ApplyNumberFormat = true
                            },
                            new CellFormat // datetime (6)
                            {
                                NumberFormatId = 22,
                                ApplyNumberFormat = true
                            },
                            new CellFormat // time (7)
                            {
                                NumberFormatId = 21,
                                ApplyNumberFormat = true
                            })
                };



                foreach (DataTable table in ds.Tables)
                {
                    var sheetPart = workbook.WorkbookPart.AddNewPart<WorksheetPart>();
                    var sheetData = new SheetData();
                    sheetPart.Worksheet = new Worksheet(sheetData);

                    Sheets sheets = workbook.WorkbookPart.Workbook.GetFirstChild<Sheets>();
                    string relationshipId = workbook.WorkbookPart.GetIdOfPart(sheetPart);

                    uint sheetId = 1;
                    if (sheets.Elements<Sheet>().Count() > 0)
                    {
                        sheetId = sheets.Elements<Sheet>().Select(s => s.SheetId.Value).Max() + 1;
                    }

                    Sheet sheet = new Sheet() { Id = relationshipId, SheetId = sheetId, Name = table.TableName };
                    sheets.Append(sheet);

                    if (columnNamesAsHeaders)
                    {
                        Row headerRow = new Row();
                        List<string> columns = new List<string>();

                        foreach (DataColumn column in table.Columns)
                        {
                            var columnName = column.ColumnName;

                            columns.Add(column.ColumnName);

                            Cell cell = new Cell();
                            cell.DataType = CellValues.String;
                            cell.CellValue = new CellValue(columnName);
                            headerRow.AppendChild(cell);
                        }

                        sheetData.AppendChild(headerRow);
                    }

                    foreach (DataRow dsrow in table.Rows)
                    {
                        Row newRow = new Row();

                        for (int col = 0; col < dsrow.ItemArray.Length; col++)
                        {
                            var value = dsrow[col];
                            var column = table.Columns[col];
                            var cellValue = value.GetType().ToCellValue();

                            Cell cell = value.ToCell();
                            //Cell cell = new Cell();
                            //cell.DataType = cellValue;
                            //cell.CellValue = cell.ToCellValue(value); // new CellValue(value.ToString());
                            newRow.AppendChild(cell);
                        }

                        sheetData.AppendChild(newRow);
                    }
                }
            }
        }
    }
}
