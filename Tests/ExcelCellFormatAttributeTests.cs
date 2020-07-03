using Flexerant.DataExport.Excel;
using System;
using System.Collections.Generic;
using System.Text;
using Xunit;

namespace Tests
{
    public class ExcelCellFormatAttributeTests
    {
        [Fact]
        public void ShortDate()
        {
            Assert.Equal("yyyy-mm-dd", ExcelCellFormatAttribute.ShortDate);
        }

        [Fact]
        public void LongDate()
        {
            Assert.Equal("[$-x-sysdate]dddd, mmmm dd, yyyy", ExcelCellFormatAttribute.LongDate);
        }

        [Fact]
        public void Currency()
        {
            Assert.Equal("$#,##0.00", ExcelCellFormatAttribute.Currency);
        }

        [Fact]
        public void Accounting()
        {
            Assert.Equal("_-$* #,##0.00_-;-$* #,##0.00_-;_-$* \"-\"??_-;_-@_-", ExcelCellFormatAttribute.Accounting);
        }

        [Fact]
        public void Percentage()
        {
            Assert.Equal("0.00%", ExcelCellFormatAttribute.Percentage);
        }

        [Fact]
        public void Number()
        {
            Assert.Equal("0.00", ExcelCellFormatAttribute.Number);
        }

        [Fact]
        public void Text()
        {
            Assert.Equal("@", ExcelCellFormatAttribute.Text);
        }
    }
}
