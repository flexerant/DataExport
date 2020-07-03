using System;
using System.Collections.Generic;
using Xunit;
using Flexerant.DataExport.Excel;

namespace Tests
{
    public class UnitTest1
    {
        [Fact]
        public void Test1()
        {
            List<Person> people = new List<Person>()
            {
                new Person()
                {
                    FirstName = "John",
                    LastName = "Smith",
                    BirthDate = new DateTime(1983, 1, 1),
                    Age = Convert.ToInt32(Math.Floor(DateTime.Now.Subtract(new DateTime(1983, 1, 1)).TotalDays / 365)),
                    IsFemale = false,
                    Worth = 1234567.89,
                    Percent = 0.1,
                    Integer = 1234
                }
            };

            ExcelWorkbook workbook = new ExcelWorkbook();

            workbook.AddSpreadheet(people);
            workbook.AddSpreadheet(people, "more people");
            workbook.Save(new System.IO.FileInfo(@"C:\Users\Kevin\Downloads\people.xlsx"));
        }
    }
}
