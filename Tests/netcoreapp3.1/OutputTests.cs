using System;
using System.Collections.Generic;
using Xunit;
using Flexerant.DataExport.Excel;
using System.IO;
using System.Data;
using ExcelDataReader;
using System.Text;
using System.Linq;

namespace Tests
{
    public class OutputTests
    {
        static OutputTests()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        }

        [Fact]
        public void Output()
        {
            List<Order> orders = new List<Order>()
            {
                new Order()
                {
                    ItemNumber = 1,
                     Description ="Party hats",
                     OrderDate = new DateTime(2019, 12, 28),
                     OrderIsComplete = false,
                     Quantity = 100,
                     UnitPrice = 0.1m,
                     SubTotal = 100 * 0.1,
                     Tax = 0.13,
                     Total = 100 * 0.1* (1 + 0.13),
                     Status = Statuses.Shipped
                },
                new Order()
                {
                    ItemNumber = 2,
                     Description ="Balloons",
                     OrderDate = new DateTime(2019, 12, 28),
                     OrderIsComplete = false,
                     Quantity = 10000,
                     UnitPrice = 0.1m,
                     SubTotal = 10000 * 0.1,
                     Tax = 0.13,
                     Total = 10000 * 0.1 * (1 + 0.13),
                     Status = Statuses.Shipped
                },
                new Order()
                {
                     ItemNumber = 3,
                     Description ="Headache medicine, extra strength, bottle of 500",
                     OrderDate = new DateTime(2020, 1, 1),
                     OrderIsComplete = false,
                     Quantity = 1,
                     UnitPrice = 15.99m,
                     SubTotal = 1 * 15.99,
                     Tax = 0.13,
                     Total = 1 * 15.99 * (1 + 0.13),
                     Status = Statuses.InProgress
                },
            };

            ExcelWorkbook workbook = new ExcelWorkbook();

            workbook.AddSpreadheet(orders);
            workbook.AddSpreadheet(orders.Where(x => x.OrderDate.Year == 2019), "2019 orders"); // Overwrite the spreadsheet name.
            workbook.AddSpreadheet(orders.Where(x => x.OrderDate.Year == 2020), "2020 orders");

            // Optionally save the workbook to the file system. To specify the output path add
            //
            // "TestSettings": {
            //     "FileOutputPath": "D:\\Users\\kevin\\Downloads\\tests\\export.xlsx"
            // }
            //
            // to the appseetings.json file in the project root. Make sure to set the file property
            // to "Copy newer" or "Copy always".

            string outputFilePath = TestHelpers.TestSettings.FileOutputPath;

            if (!string.IsNullOrWhiteSpace(outputFilePath))
            {
                FileInfo fi = new FileInfo(outputFilePath);

                if (fi.Exists) fi.Delete();

                workbook.Save(fi);
            }

            DataSet ds;
            byte[] excelData;

            using (MemoryStream ms = new MemoryStream())
            {
                workbook.Save(ms);
                excelData = ms.ToArray();
            }

            using (MemoryStream ms = new MemoryStream(excelData))
            {
                using (var reader = ExcelReaderFactory.CreateReader(ms))
                {
                    ds = reader.AsDataSet();
                }
            }

            // Confirm the expected spreadsheets exist.
            Assert.NotNull(ds.Tables["Orders"]);
            Assert.NotNull(ds.Tables["2019 orders"]);
            Assert.NotNull(ds.Tables["2020 orders"]);

            var headingRow = ds.Tables["Orders"].Rows[0];

            // Confirm the order of columns is as expected.
            Assert.Equal("Product description", headingRow.Field<string>(1));
            Assert.Equal("Order date", headingRow.Field<string>(2));
            Assert.Equal("Quantity", headingRow.Field<string>(3));
            Assert.Equal("Price", headingRow.Field<string>(4));
            Assert.Equal("Sub-total", headingRow.Field<string>(5));
            Assert.Equal("Tax", headingRow.Field<string>(6));
            Assert.Equal("Total", headingRow.Field<string>(7));
            Assert.Equal("Order is complete", headingRow.Field<string>(9));

            // Confirm the row counts
            Assert.Equal(4, ds.Tables["Orders"].Rows.Count);
            Assert.Equal(3, ds.Tables["2019 orders"].Rows.Count);
            Assert.Equal(2, ds.Tables["2020 orders"].Rows.Count);

            // Confirm only the non-ignored column headings are displayed.
            Assert.Equal(typeof(Order).GetProperties().Length - 1, ds.Tables["Orders"].Rows[0].ItemArray.Length);
            Assert.Equal(typeof(Order).GetProperties().Length - 1, ds.Tables["2019 orders"].Rows[0].ItemArray.Length);
            Assert.Equal(typeof(Order).GetProperties().Length - 1, ds.Tables["2020 orders"].Rows[0].ItemArray.Length);
        }
    }
}
