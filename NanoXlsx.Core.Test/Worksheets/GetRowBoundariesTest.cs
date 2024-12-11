using NanoXLSX;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xunit;
using static NanoXLSX.Test.Worksheets.RowTest;

namespace NanoXLSX.Test.Worksheets
{
    public class GetRowBoundariesTest
    {
     
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

        [Theory(DisplayName = "Test of the GetLastRowNumber function with an explicitly defined, empty cell besides other row definitions")]
        [InlineData("F7", 6)]
        [InlineData("A1", 4)]
        public void GetLastRowNumberTest6(string emptyCellAddress, int expectedFirstRow)
        {
            Worksheet worksheet = new Worksheet();
            worksheet.AddHiddenRow(3);
            worksheet.AddHiddenRow(4);
            worksheet.AddCell(null, emptyCellAddress);
            int row = worksheet.GetLastRowNumber();
            Assert.Equal(expectedFirstRow, row);
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

        [Fact(DisplayName = "Test of the GetFirstRowNumber function with an empty worksheet")]
        public void GetFirstRowNumberTest()
        {
            Worksheet worksheet = new Worksheet();
            int row = worksheet.GetFirstRowNumber();
            Assert.Equal(-1, row);
        }

        [Theory(DisplayName = "Test of the GetFirstRowNumber function with defined rows on an empty worksheet")]
        [InlineData(RowProperty.Height)]
        [InlineData(RowProperty.Hidden)]
        public void GetFisrtRowNumberTest2(RowProperty rowProperty)
        {
            Worksheet worksheet = new Worksheet();
            if (rowProperty == RowProperty.Hidden)
            {
                worksheet.AddHiddenRow(1);
                worksheet.AddHiddenRow(2);
                worksheet.AddHiddenRow(3);
            }
            else
            {
                worksheet.SetRowHeight(1, 22.2f);
                worksheet.SetRowHeight(2, 33.3f);
                worksheet.SetRowHeight(3, 44.4f);
            }
            int row = worksheet.GetFirstRowNumber();
            Assert.Equal(1, row);
        }

        [Theory(DisplayName = "Test of the GetFirstRowNumber function with defined rows on an empty worksheet, where the row definition has gaps")]
        [InlineData(RowProperty.Height)]
        [InlineData(RowProperty.Hidden)]
        public void GetFirstRowNumberTest3(RowProperty rowProperty)
        {
            Worksheet worksheet = new Worksheet();
            if (rowProperty == RowProperty.Hidden)
            {
                worksheet.AddHiddenRow(1);
                worksheet.AddHiddenRow(2);
                worksheet.AddHiddenRow(10);
            }
            else
            {
                worksheet.SetRowHeight(1, 22.2f);
                worksheet.SetRowHeight(2, 33.3f);
                worksheet.SetRowHeight(10, 44.4f);
            }
            int row = worksheet.GetFirstRowNumber();
            Assert.Equal(1, row);
        }

        [Theory(DisplayName = "Test of the GetFirstRowNumber function with defined rows where cells are defined above the first row")]
        [InlineData(RowProperty.Height)]
        [InlineData(RowProperty.Hidden)]
        public void GetFirstRowNumberTest4(RowProperty rowProperty)
        {
            Worksheet worksheet = new Worksheet();
            if (rowProperty == RowProperty.Hidden)
            {
                worksheet.AddHiddenRow(2);
                worksheet.AddHiddenRow(3);
                worksheet.AddHiddenRow(10);
            }
            else
            {
                worksheet.SetRowHeight(2, 22.2f);
                worksheet.SetRowHeight(3, 33.3f);
                worksheet.SetRowHeight(10, 44.4f);
            }
            worksheet.AddCell("test", "E5");
            int row = worksheet.GetFirstRowNumber();
            Assert.Equal(2, row);
        }

        [Theory(DisplayName = "Test of the GetFirstRowNumber function with defined rows where cells are defined below the first row")]
        [InlineData(RowProperty.Height)]
        [InlineData(RowProperty.Hidden)]
        public void GetFirstRowNumberTest5(RowProperty rowProperty)
        {
            Worksheet worksheet = new Worksheet();
            if (rowProperty == RowProperty.Hidden)
            {
                worksheet.AddHiddenRow(6);
                worksheet.AddHiddenRow(7);
                worksheet.AddHiddenRow(8);
            }
            else
            {
                worksheet.SetRowHeight(6, 22.2f);
                worksheet.SetRowHeight(7, 33.3f);
                worksheet.SetRowHeight(8, 44.4f);
            }
            worksheet.AddCell("test", "F5");
            int row = worksheet.GetFirstRowNumber();
            Assert.Equal(4, row);
        }

        [Theory(DisplayName = "Test of the getFirstRowNumber function with an explicitly defined, empty cell besides other row definitions")]
        [InlineData("F5", 4)]
        [InlineData("A1", 0)]
        public void GetFirstRowNumberTest6(string emptyCellAddress, int expectedFirstRow)
        {
            Worksheet worksheet = new Worksheet();
            worksheet.AddHiddenColumn(3);
            worksheet.AddHiddenColumn(4);
            worksheet.AddCell(null, emptyCellAddress);
            int row = worksheet.GetFirstRowNumber();
            Assert.Equal(expectedFirstRow, row);
        }

        [Fact(DisplayName = "Test of the GetFirstDataRowNumber function with an empty worksheet")]
        public void GetFirstDataRowNumberTest()
        {
            Worksheet worksheet = new Worksheet();
            int row = worksheet.GetFirstDataRowNumber();
            Assert.Equal(-1, row);
        }

        [Theory(DisplayName = "Test of the GetFirstDataRowNumber function with defined rows on an empty worksheet")]
        [InlineData(RowProperty.Height)]
        [InlineData(RowProperty.Hidden)]
        public void GetFirstDataRowNumberTest2(RowProperty rowProperty)
        {
            Worksheet worksheet = new Worksheet();
            if (rowProperty == RowProperty.Hidden)
            {
                worksheet.AddHiddenRow(1);
                worksheet.AddHiddenRow(2);
                worksheet.AddHiddenRow(3);
            }
            else
            {
                worksheet.SetRowHeight(1, 22.2f);
                worksheet.SetRowHeight(2, 33.3f);
                worksheet.SetRowHeight(3, 44.4f);
            }
            int row = worksheet.GetFirstDataRowNumber();
            Assert.Equal(-1, row);
        }

        [Theory(DisplayName = "Test of the GetFirstDataRowNumber function with defined rows where cells are defined below the last row")]
        [InlineData(RowProperty.Height)]
        [InlineData(RowProperty.Hidden)]
        public void GetFirstDataRowNumberTest3(RowProperty rowProperty)
        {
            Worksheet worksheet = new Worksheet();
            if (rowProperty == RowProperty.Hidden)
            {
                worksheet.AddHiddenRow(2);
                worksheet.AddHiddenRow(3);
                worksheet.AddHiddenRow(10);
            }
            else
            {
                worksheet.SetRowHeight(2, 22.2f);
                worksheet.SetRowHeight(3, 33.3f);
                worksheet.SetRowHeight(10, 44.4f);
            }

            worksheet.AddCell("test", "E5");
            int row = worksheet.GetFirstDataRowNumber();
            Assert.Equal(4, row);
        }

        [Theory(DisplayName = "Test of the GetFirstDataRowNumber function with defined rows where cells are defined above the last row")]
        [InlineData(RowProperty.Height)]
        [InlineData(RowProperty.Hidden)]
        public void GetfirstDataRowNumberTest4(RowProperty rowProperty)
        {
            Worksheet worksheet = new Worksheet();
            if (rowProperty == RowProperty.Hidden)
            {
                worksheet.AddHiddenRow(1);
                worksheet.AddHiddenRow(2);
                worksheet.AddHiddenRow(3);
            }
            else
            {
                worksheet.SetRowHeight(1, 22.2f);
                worksheet.SetRowHeight(2, 33.3f);
                worksheet.SetRowHeight(3, 44.4f);
            }

            worksheet.AddCell("test", "F6");
            int row = worksheet.GetFirstDataRowNumber();
            Assert.Equal(5, row);
        }

        [Theory(DisplayName = "Test of the getFirstDataRowNumber and getLastDataRowNumber functions with an explicitly defined, empty cell besides other row definitions")]
        [InlineData("F5")]
        [InlineData("A1")]
        void GetFirstOrLastDataRowNumberTest(string emptyCellAddress)
        {
            Worksheet worksheet = new Worksheet();
            worksheet.AddHiddenRow(3);
            worksheet.AddHiddenRow(4);
            worksheet.AddCell(null, emptyCellAddress);
            int minRow = worksheet.GetFirstDataRowNumber();
            int maxRow = worksheet.GetLastDataRowNumber();
            Assert.Equal(-1, minRow);
            Assert.Equal(-1, maxRow);
        }

        [Fact(DisplayName = "Test of the GetFirstDataRowNumber and GetLastDataRowNumber functions with exactly one defined cell")]
        public void GetFirstOrLastDataRowNumberTest2()
        {
            Worksheet worksheet = new Worksheet();
            worksheet.AddHiddenRow(2);
            worksheet.AddHiddenRow(3);
            worksheet.AddHiddenRow(10);
            worksheet.AddCell("test", "C5");
            int minRow = worksheet.GetFirstDataRowNumber();
            int maxRow = worksheet.GetLastDataRowNumber();
            Assert.Equal(4, minRow);
            Assert.Equal(4, maxRow);
        }

        [Theory(DisplayName = "Test of the GetFirstDataRowNumber and GetLastDataRowNumber functions with an explicitly defined, cell with empty string besides other row definitions")]
        [InlineData("F5")]
        [InlineData("A1")]
        void GetFirstOrLastDataRowNumberTest3(string emptyCellAddress)
        {
            Worksheet worksheet = new Worksheet();
            worksheet.AddHiddenRow(3);
            worksheet.AddHiddenRow(4);
            worksheet.AddCell(string.Empty, emptyCellAddress);
            int minRow = worksheet.GetFirstDataRowNumber();
            int maxRow = worksheet.GetLastDataRowNumber();
            Assert.Equal(-1, minRow);
            Assert.Equal(-1, maxRow);
        }
    }
}
