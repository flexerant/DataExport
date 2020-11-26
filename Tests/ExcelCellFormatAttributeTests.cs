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
            Assert.Equal("yyyy-mm-dd", CellFormats.SHORT_DATE);
        }

        [Fact]
        public void LongDate()
        {
            Assert.Equal("[$-x-sysdate]dddd, mmmm dd, yyyy", CellFormats.LONG_DATE);
        }

        [Fact]
        public void Currency()
        {
            Assert.Equal("$#,##0.00", CellFormats.CURRENCY);
        }

        [Fact]
        public void Accounting()
        {
            Assert.Equal("_-$* #,##0.00_-;-$* #,##0.00_-;_-$* \"-\"??_-;_-@_-", CellFormats.ACCOUNTING);
        }

        [Fact]
        public void Percentage()
        {
            Assert.Equal("0.00%", CellFormats.PERCENTAGE);
        }

        [Fact]
        public void Number()
        {
            Assert.Equal("0.00", CellFormats.NUMBER);
        }

        [Fact]
        public void Text()
        {
            Assert.Equal("@", CellFormats.TEXT);
        }
    }
}
