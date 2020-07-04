using Flexerant.DataExport.Excel;
using System;
using System.Collections.Generic;
using System.Security.Cryptography;
using System.Text;

namespace Tests
{
    [ExcelSpreadsheet("Orders")]
    public class Order
    {
        [ExcelSpreadsheetIgnoreColumn()]
        public Guid OrderId { get; set; } = new Guid();

        [ExcelSpreadsheetColumn("Product description", Order = 0)]
        public string Description { get; set; }

        [ExcelSpreadsheetColumn("Order date", Order = 1)]
        [ExcelCellFormat(ExcelCellFormatAttribute.ShortDate)]
        public DateTime OrderDate { get; set; }

        [ExcelSpreadsheetColumn(Order = 2)]
        public int Quantity { get; set; }

        [ExcelSpreadsheetColumn("Order is complete")]
        public bool OrderIsComplete { get; set; }

        [ExcelSpreadsheetColumn("Price", Order = 3)]
        [ExcelCellFormat(ExcelCellFormatAttribute.Accounting)]
        public decimal UnitPrice { get; set; }

        [ExcelSpreadsheetColumn("Sub-total", Order = 4)]
        [ExcelCellFormat(ExcelCellFormatAttribute.Accounting)]
        public double SubTotal { get; set; }

        [ExcelSpreadsheetColumn(Order = 6)]
        [ExcelCellFormat(ExcelCellFormatAttribute.Accounting)]
        public double Total { get; set; }

        [ExcelSpreadsheetColumn(Order = 5)]
        [ExcelCellFormat(ExcelCellFormatAttribute.Percentage)]
        public double Tax { get; set; }
    }
}
