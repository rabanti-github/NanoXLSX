using NanoXLSX;
using NanoXLSX.Exceptions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xunit;

namespace NanoXLSX_Test.Worksheets
{
    public class RowTest
    {
        public enum RowProperty
        {
            Hidden,
            Height
        }

        [Fact(DisplayName = "Test of the AddHiddenRow function with a row number")]
        public void AddHiddenRowTest()
        {
            Worksheet worksheet = new Worksheet();
            Assert.Empty(worksheet.HiddenRows);
            worksheet.AddHiddenRow(2);
            Assert.Single(worksheet.HiddenRows);
            Assert.Contains(worksheet.HiddenRows, item => item.Key == 2);
            Assert.True(worksheet.HiddenRows[2]);
            worksheet.AddHiddenRow(2); // Should not add an additional entry
            Assert.Single(worksheet.HiddenRows);
        }

        [Theory(DisplayName = "Test of the failing AddHiddenRow function with an invalid row number")]
        [InlineData(-1)]
        [InlineData(-100)]
        [InlineData(1048576)]
        public void AddHiddenRowFailTest(int value)
        {
            Worksheet worksheet = new Worksheet();
            Assert.Throws<RangeException>(() => worksheet.AddHiddenRow(value));
        }

        [Fact(DisplayName = "Test of the GetLastRowNumber function with an empty worksheet")]
        public void GetLastRowNumberTest()
        {
            Worksheet worksheet = new Worksheet();
            int row = worksheet.GetLastRowNumber();
            Assert.Equal(-1, row);
        }

        [Theory(DisplayName = "Test of the GetLastRowNumber function with defined rows on an empty worksheet")]
        [InlineData(RowProperty.Height)]
        [InlineData(RowProperty.Hidden)]
        public void GetLastRowNumberTest2(RowProperty rowProperty)
        {
            Worksheet worksheet = new Worksheet();
            if (rowProperty == RowProperty.Hidden)
            {
                worksheet.AddHiddenRow(0);
                worksheet.AddHiddenRow(1);
                worksheet.AddHiddenRow(2);
            }
            else
            {
                worksheet.SetRowHeight(0, 22.2f);
                worksheet.SetRowHeight(1, 33.3f);
                worksheet.SetRowHeight(2, 44.4f);
            }
            int row = worksheet.GetLastRowNumber();
            Assert.Equal(2, row);
        }

        [Theory(DisplayName = "Test of the GetLastRowNumber function with defined rows on an empty worksheet, where the row definition has gaps")]
        [InlineData(RowProperty.Height)]
        [InlineData(RowProperty.Hidden)]
        public void GetLastRowNumberTest3(RowProperty rowProperty)
        {
            Worksheet worksheet = new Worksheet();
            if (rowProperty == RowProperty.Hidden)
            {
                worksheet.AddHiddenRow(0);
                worksheet.AddHiddenRow(1);
                worksheet.AddHiddenRow(10);
            }
            else
            {
                worksheet.SetRowHeight(0, 22.2f);
                worksheet.SetRowHeight(1, 33.3f);
                worksheet.SetRowHeight(10, 44.4f);
            }
            int row = worksheet.GetLastRowNumber();
            Assert.Equal(10, row);
        }

        [Theory(DisplayName = "Test of the GetLastRowNumber function with defined rows where cells are defined below the last row")]
        [InlineData(RowProperty.Height)]
        [InlineData(RowProperty.Hidden)]
        public void GetLastRowNumberTest4(RowProperty rowProperty)
        {
            Worksheet worksheet = new Worksheet();
            if (rowProperty == RowProperty.Hidden)
            {
                worksheet.AddHiddenRow(0);
                worksheet.AddHiddenRow(1);
                worksheet.AddHiddenRow(10);
            }
            else
            {
                worksheet.SetRowHeight(0, 22.2f);
                worksheet.SetRowHeight(1, 33.3f);
                worksheet.SetRowHeight(10, 44.4f);
            }
            worksheet.AddCell("test", "E5");
            int row = worksheet.GetLastRowNumber();
            Assert.Equal(10, row);
        }

        [Theory(DisplayName = "Test of the GetLastRowNumber function with defined rows where cells are defined above the last row")]
        [InlineData(RowProperty.Height)]
        [InlineData(RowProperty.Hidden)]
        public void GetLastRowNumberTest5(RowProperty rowProperty)
        {
            Worksheet worksheet = new Worksheet();
            if (rowProperty == RowProperty.Hidden)
            {
                worksheet.AddHiddenRow(0);
                worksheet.AddHiddenRow(1);
                worksheet.AddHiddenRow(2);
            }
            else
            {
                worksheet.SetRowHeight(0, 22.2f);
                worksheet.SetRowHeight(1, 33.3f);
                worksheet.SetRowHeight(2, 44.4f);
            }
            worksheet.AddCell("test", "F5");
            int row = worksheet.GetLastRowNumber();
            Assert.Equal(4, row);
        }

        [Fact(DisplayName = "Test of the GetLastDataRowNumber function with an empty worksheet")]
        public void GetLastDataRowNumberTest()
        {
            Worksheet worksheet = new Worksheet();
            int row = worksheet.GetLastDataRowNumber();
            Assert.Equal(-1, row);
        }

        [Theory(DisplayName = "Test of the GetLastDataRowNumber function with defined rows on an empty worksheet")]
        [InlineData(RowProperty.Height)]
        [InlineData(RowProperty.Hidden)]
        public void GetLastDataRowNumberTest2(RowProperty rowProperty)
        {
            Worksheet worksheet = new Worksheet();
            if (rowProperty == RowProperty.Hidden)
            {
                worksheet.AddHiddenRow(0);
                worksheet.AddHiddenRow(1);
                worksheet.AddHiddenRow(2);
            }
            else
            {
                worksheet.SetRowHeight(0, 22.2f);
                worksheet.SetRowHeight(1, 33.3f);
                worksheet.SetRowHeight(2, 44.4f);
            }
            int row = worksheet.GetLastDataRowNumber();
            Assert.Equal(-1, row);
        }

        [Theory(DisplayName = "Test of the GetLastDataRowNumber function with defined rows where cells are defined below the last row")]
        [InlineData(RowProperty.Height)]
        [InlineData(RowProperty.Hidden)]
        public void GetLastDataRowNumberTest3(RowProperty rowProperty)
        {
            Worksheet worksheet = new Worksheet();
            if (rowProperty == RowProperty.Hidden)
            {
                worksheet.AddHiddenRow(0);
                worksheet.AddHiddenRow(1);
                worksheet.AddHiddenRow(10);
            }
            else
            {
                worksheet.SetRowHeight(0, 22.2f);
                worksheet.SetRowHeight(1, 33.3f);
                worksheet.SetRowHeight(10, 44.4f);
            }

            worksheet.AddCell("test", "E5");
            int row = worksheet.GetLastDataRowNumber();
            Assert.Equal(4, row);
        }

        [Theory(DisplayName = "Test of the GetLastDataRowNumber function with defined rows where cells are defined above the last row")]
        [InlineData(RowProperty.Height)]
        [InlineData(RowProperty.Hidden)]
        public void GetLastDataRowNumberTest4(RowProperty rowProperty)
        {
            Worksheet worksheet = new Worksheet();
            if (rowProperty == RowProperty.Hidden)
            {
                worksheet.AddHiddenRow(0);
                worksheet.AddHiddenRow(1);
                worksheet.AddHiddenRow(2);
            }
            else
            {
                worksheet.SetRowHeight(0, 22.2f);
                worksheet.SetRowHeight(1, 33.3f);
                worksheet.SetRowHeight(3, 44.4f);
            }

            worksheet.AddCell("test", "F5");
            int row = worksheet.GetLastDataRowNumber();
            Assert.Equal(4, row);
        }






    }
}
