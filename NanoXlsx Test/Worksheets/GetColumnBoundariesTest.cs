using NanoXLSX;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xunit;

namespace NanoXLSX_Test.Worksheets
{
    public class GetColumnBoundariesTest
    {
        [Fact(DisplayName = "Test of the GetLastColumnNumber function with an empty worksheet")]
        public void GetLastColumnNumberTest()
        {
            Worksheet worksheet = new Worksheet();
            int column = worksheet.GetLastColumnNumber();
            Assert.Equal(-1, column);
        }

        [Fact(DisplayName = "Test of the GetLastColumnNumber function with defined columns on an empty worksheet")]
        public void GetLastColumnNumberTest2()
        {
            Worksheet worksheet = new Worksheet();
            worksheet.AddHiddenColumn(0);
            worksheet.AddHiddenColumn(1);
            worksheet.AddHiddenColumn(2);
            int column = worksheet.GetLastColumnNumber();
            Assert.Equal(2, column);
        }

        [Fact(DisplayName = "Test of the GetLastColumnNumber function with defined columns on an empty worksheet, where the column definition has gaps")]
        public void GetLastColumnNumberTest3()
        {
            Worksheet worksheet = new Worksheet();
            worksheet.AddHiddenColumn(0);
            worksheet.AddHiddenColumn(1);
            worksheet.AddHiddenColumn(10);
            int column = worksheet.GetLastColumnNumber();
            Assert.Equal(10, column);
        }

        [Fact(DisplayName = "Test of the GetLastColumnNumber function with defined columns where cells are defined below the last column")]
        public void GetLastColumnNumberTest4()
        {
            Worksheet worksheet = new Worksheet();
            worksheet.AddHiddenColumn(0);
            worksheet.AddHiddenColumn(1);
            worksheet.AddHiddenColumn(10);
            worksheet.AddCell("test", "E5");
            int column = worksheet.GetLastColumnNumber();
            Assert.Equal(10, column);
        }

        [Fact(DisplayName = "Test of the GetLastColumnNumber function with defined columns where cells are defined above the last column")]
        public void GetLastColumnNumberTest5()
        {
            Worksheet worksheet = new Worksheet();
            worksheet.AddHiddenColumn(0);
            worksheet.AddHiddenColumn(1);
            worksheet.AddHiddenColumn(2);
            worksheet.AddCell("test", "F5");
            int column = worksheet.GetLastColumnNumber();
            Assert.Equal(5, column);
        }


        [Fact(DisplayName = "Test of the GetFirstColumnNumber function with an empty worksheet")]
        public void GetFirstColumnNumberTest()
        {
            Worksheet worksheet = new Worksheet();
            int column = worksheet.GetFirstColumnNumber();
            Assert.Equal(-1, column);
        }

        [Fact(DisplayName = "Test of the GetFirstColumnNumber function with defined columns on an empty worksheet")]
        public void GetFirstColumnNumberTest2()
        {
            Worksheet worksheet = new Worksheet();
            worksheet.AddHiddenColumn(1);
            worksheet.AddHiddenColumn(2);
            worksheet.AddHiddenColumn(3);
            int column = worksheet.GetFirstColumnNumber();
            Assert.Equal(1, column);
        }

        [Fact(DisplayName = "Test of the GetFirstColumnNumber function with defined columns on an empty worksheet, where the column definition has gaps")]
        public void GetFirstColumnNumberTest3()
        {
            Worksheet worksheet = new Worksheet();
            worksheet.AddHiddenColumn(1);
            worksheet.AddHiddenColumn(2);
            worksheet.AddHiddenColumn(10);
            int column = worksheet.GetFirstColumnNumber();
            Assert.Equal(1, column);
        }

        [Fact(DisplayName = "Test of the GetFirstColumnNumber function with defined columns where cells are defined above the first column")]
        public void GetFirstColumnNumberTest4()
        {
            Worksheet worksheet = new Worksheet();
            worksheet.AddHiddenColumn(3);
            worksheet.AddHiddenColumn(8);
            worksheet.AddHiddenColumn(10);
            worksheet.AddCell("test", "E5");
            int column = worksheet.GetFirstColumnNumber();
            Assert.Equal(3, column);
        }

        [Fact(DisplayName = "Test of the GetFirstColumnNumber function with defined columns where cells are defined below the first column")]
        public void GetFirstColumnNumberTest5()
        {
            Worksheet worksheet = new Worksheet();
            worksheet.AddHiddenColumn(7);
            worksheet.AddHiddenColumn(8);
            worksheet.AddHiddenColumn(9);
            worksheet.AddCell("test", "F5");
            int column = worksheet.GetFirstColumnNumber();
            Assert.Equal(5, column);
        }

        [Fact(DisplayName = "Test of the GetLastDataColumnNumber function with an empty worksheet")]
        public void GetLastDataColumnNumberTest()
        {
            Worksheet worksheet = new Worksheet();
            int column = worksheet.GetLastDataColumnNumber();
            Assert.Equal(-1, column);
        }

        [Fact(DisplayName = "Test of the GetLastDataColumnNumber function with defined columns on an empty worksheet")]
        public void GetLastDataColumnNumberTest2()
        {
            Worksheet worksheet = new Worksheet();
            worksheet.AddHiddenColumn(0);
            worksheet.AddHiddenColumn(1);
            worksheet.AddHiddenColumn(2);
            int column = worksheet.GetLastDataColumnNumber();
            Assert.Equal(-1, column);
        }

        [Fact(DisplayName = "Test of the GetLastDataColumnNumber function with defined columns where cells are defined below the last column")]
        public void GetLastDataColumnNumberTest3()
        {
            Worksheet worksheet = new Worksheet();
            worksheet.AddHiddenColumn(0);
            worksheet.AddHiddenColumn(1);
            worksheet.AddHiddenColumn(10);
            worksheet.AddCell("test", "E5");
            int column = worksheet.GetLastDataColumnNumber();
            Assert.Equal(4, column);
        }

        [Fact(DisplayName = "Test of the GetLastDataColumnNumber function with defined columns where cells are defined above the last column")]
        public void GetLastDataColumnNumberTest4()
        {
            Worksheet worksheet = new Worksheet();
            worksheet.AddHiddenColumn(0);
            worksheet.AddHiddenColumn(1);
            worksheet.AddHiddenColumn(2);
            worksheet.AddCell("test", "F5");
            int column = worksheet.GetLastDataColumnNumber();
            Assert.Equal(5, column);
        }

        [Fact(DisplayName = "Test of the GetFirstDataColumnNumber function with an empty worksheet")]
        public void GetFirstDataColumnNumberTest()
        {
            Worksheet worksheet = new Worksheet();
            int column = worksheet.GetFirstDataColumnNumber();
            Assert.Equal(-1, column);
        }

        [Fact(DisplayName = "Test of the GetFirstDataColumnNumber function with defined columns on an empty worksheet")]
        public void GetFirstDataColumnNumberTest2()
        {
            Worksheet worksheet = new Worksheet();
            worksheet.AddHiddenColumn(0);
            worksheet.AddHiddenColumn(1);
            worksheet.AddHiddenColumn(2);
            int column = worksheet.GetFirstDataColumnNumber();
            Assert.Equal(-1, column);
        }

        [Fact(DisplayName = "Test of the GetFirstDataColumnNumber function with defined columns where cells are defined above the first column")]
        public void GetFirstDataColumnNumberTest3()
        {
            Worksheet worksheet = new Worksheet();
            worksheet.AddHiddenColumn(2);
            worksheet.AddHiddenColumn(3);
            worksheet.AddHiddenColumn(10);
            worksheet.AddCell("test", "E5");
            int column = worksheet.GetFirstDataColumnNumber();
            Assert.Equal(4, column);
        }

        [Fact(DisplayName = "Test of the GetFirstDataColumnNumber function with defined columns where cells are defined below the first column")]
        public void GetFirstDataColumnNumberTest4()
        {
            Worksheet worksheet = new Worksheet();
            worksheet.AddHiddenColumn(2);
            worksheet.AddHiddenColumn(3);
            worksheet.AddHiddenColumn(10);
            worksheet.AddCell("test", "F5");
            int column = worksheet.GetFirstDataColumnNumber();
            Assert.Equal(5, column);
        }
    }
}
