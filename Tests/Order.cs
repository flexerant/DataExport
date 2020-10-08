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

        [ExcelSpreadsheetColumn("Item number", Order = 0)]
        public int ItemNumber { get; set; }

        [ExcelSpreadsheetColumn("Product description", Order = 1)]
        public string Description { get; set; }

        [ExcelSpreadsheetColumn("Order date", Order = 2, CellFormat = CellFormats.SHORT_DATE)]
        public DateTime OrderDate { get; set; }

        [ExcelSpreadsheetColumn(Order = 3, CellFormat = "_-* #,##0_-;-* #,##0_-;_-* \"-\"??_-;_-@_-")]
        public int Quantity { get; set; }

        [ExcelSpreadsheetColumn("Order is complete")]
        public bool OrderIsComplete { get; set; }

        [ExcelSpreadsheetColumn("Price", Order = 4, CellFormat = CellFormats.ACCOUNTING)]
        public decimal UnitPrice { get; set; }

        [ExcelSpreadsheetColumn("Sub-total", Order = 5, CellFormat = CellFormats.ACCOUNTING)]        
        public double SubTotal { get; set; }

        [ExcelSpreadsheetColumn(Order = 7, CellFormat = CellFormats.ACCOUNTING)]        
        public double Total { get; set; }

        [ExcelSpreadsheetColumn(Order = 6, CellFormat = CellFormats.PERCENTAGE)]
        public double Tax { get; set; }

        [ExcelSpreadsheetColumn("Order status", Order = 8)]
        public Statuses Status { get; set; }
    }
}
