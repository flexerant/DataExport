using System;
using System.Collections.Generic;
using Xunit;
using Flexerant.DataExport.Excel;
using System.IO;
using System.Data;
using ExcelDataReader;
using System.Reflection.Emit;
using Newtonsoft.Json;
using System.Text;

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
            string json1 = Encoding.UTF8.GetString(Properties.Resources.MOCK_PERSON_DATA_1, 0, Properties.Resources.MOCK_PERSON_DATA_1.Length);
            string json2 = Encoding.UTF8.GetString(Properties.Resources.MOCK_PERSON_DATA_2, 0, Properties.Resources.MOCK_PERSON_DATA_2.Length);
            JsonSerializerSettings serializerSettings = new JsonSerializerSettings { DateFormatString = "MM/dd/yyyy" };
            List<Person> people1 = JsonConvert.DeserializeObject<List<Person>>(json1, serializerSettings);
            List<Person> people2 = JsonConvert.DeserializeObject<List<Person>>(json2, serializerSettings);
            ExcelWorkbook workbook = new ExcelWorkbook();

            workbook.AddSpreadheet(people1);
            workbook.AddSpreadheet(people2, "more people"); // override the sheet name.

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
            Assert.NotNull(ds.Tables["people"]);
            Assert.NotNull(ds.Tables["more people"]);

            var headingRow = ds.Tables["people"].Rows[0];

            // Confirm the order of columns is as expected.
            Assert.Equal("First name", headingRow.Field<string>(0));
            Assert.Equal("Last name", headingRow.Field<string>(1));
            Assert.Equal("Date of birth", headingRow.Field<string>(2));
            Assert.Equal("Age", headingRow.Field<string>(3));
            Assert.Equal("Female", headingRow.Field<string>(4));
            Assert.Equal("Integer", headingRow.Field<string>(5));
            Assert.Equal("Percent", headingRow.Field<string>(6));
            Assert.Equal("Text", headingRow.Field<string>(7));
            Assert.Equal("Worth", headingRow.Field<string>(8));

            // Confirm the row counts
            Assert.Equal(1001, ds.Tables["people"].Rows.Count);
            Assert.Equal(1001, ds.Tables["more people"].Rows.Count);

            // Confirm only the non-ignored column headings are displayed.
            Assert.Equal(9, ds.Tables["people"].Rows[0].ItemArray.Length);
            Assert.Equal(9, ds.Tables["more people"].Rows[0].ItemArray.Length);

            // Confirm only the non-ignored columns are displayed.
            Assert.Equal(9, ds.Tables["people"].Rows[10].ItemArray.Length);
            Assert.Equal(9, ds.Tables["people"].Rows[100].ItemArray.Length);
            Assert.Equal(9, ds.Tables["people"].Rows[1000].ItemArray.Length);
            Assert.Equal(9, ds.Tables["more people"].Rows[10].ItemArray.Length);
            Assert.Equal(9, ds.Tables["more people"].Rows[100].ItemArray.Length);
            Assert.Equal(9, ds.Tables["more people"].Rows[1000].ItemArray.Length);
        }
    }
}
