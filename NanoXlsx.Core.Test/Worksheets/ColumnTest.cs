using NanoXLSX;
using NanoXLSX.Shared.Exceptions;
using NanoXLSX.Styles;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xunit;

namespace NanoXLSX.Test.Worksheets
{
    public class ColumnTest
    {
        [Fact(DisplayName = "Test of the AddHiddenColumn function with a column number")]
        public void AddHiddenColumnTest()
        {
            Worksheet worksheet = new Worksheet();
            Assert.Empty(worksheet.Columns);
            worksheet.AddHiddenColumn(2);
            Assert.Single(worksheet.Columns);
            Assert.Contains(worksheet.Columns, item => item.Key == 2);
            Assert.True(worksheet.Columns[2].IsHidden);
            worksheet.AddHiddenColumn(2); // Should not add an additional entry
            Assert.Single(worksheet.Columns);
        }

        [Fact(DisplayName = "Test of the AddHiddenColumn function with a column as string")]
        public void AddHiddenColumnTest2()
        {
            Worksheet worksheet = new Worksheet();
            Assert.Empty(worksheet.Columns);
            worksheet.AddHiddenColumn("C");
            Assert.Single(worksheet.Columns);
            Assert.Contains(worksheet.Columns, item => item.Key == 2);
            Assert.True(worksheet.Columns[2].IsHidden);
            worksheet.AddHiddenColumn("C"); // Should not add an additional entry
            Assert.Single(worksheet.Columns);
        }

        [Theory(DisplayName = "Test of the failing AddHiddenColumn function with an invalid column number")]
        [InlineData(-1)]
        [InlineData(-100)]
        [InlineData(16384)]
        public void AddHiddenColumnFailTest(int value)
        {
            Worksheet worksheet = new Worksheet();
            Assert.Throws<RangeException>(() => worksheet.AddHiddenColumn(value));
        }

        [Theory(DisplayName = "Test of the failing AddHiddenColumn function with an invalid column string")]
        [InlineData(null)]
        [InlineData("")]
        [InlineData("#")]
        [InlineData("XFE")]
        public void AddHiddenColumnFailTest2(string value)
        {
            Worksheet worksheet = new Worksheet();
            Assert.Throws<RangeException>(() => worksheet.AddHiddenColumn(value));
        }

        [Fact(DisplayName = "Test of the ResetColumn function with an empty worksheet")]
        public void ResetColumnTest()
        {
            Worksheet worksheet = new Worksheet();
            Assert.Empty(worksheet.Columns);
            worksheet.ResetColumn(0); // Should do nothing and not fail
            Assert.Empty(worksheet.Columns);
        }

        [Fact(DisplayName = "Test of the ResetColumn function with defined columns but a not defined columns")]
        public void ResetColumnTest2()
        {
            Worksheet worksheet = new Worksheet();
            Assert.Empty(worksheet.Columns);
            worksheet.AddHiddenColumn(0);
            worksheet.AddHiddenColumn(2);
            worksheet.ResetColumn(1); // Should do nothing and not fail
            Assert.Equal(2, worksheet.Columns.Count);
        }

        [Fact(DisplayName = "Test of the ResetColumn function with defined columns and an existing column")]
        public void ResetColumnTest3()
        {
            Worksheet worksheet = new Worksheet();
            Assert.Empty(worksheet.Columns);
            worksheet.AddHiddenColumn(0);
            worksheet.AddHiddenColumn(1);
            worksheet.AddHiddenColumn(2);
            Assert.Equal(3, worksheet.Columns.Count);
            worksheet.ResetColumn(1);
            Assert.Equal(2, worksheet.Columns.Count);
            Assert.DoesNotContain(worksheet.Columns, item => item.Key == 1);
        }

        [Fact(DisplayName = "Test of the ResetColumn function with defined columns and a AutoFilter definition")]
        public void ResetColumnTest4()
        {
            Worksheet worksheet = new Worksheet();
            Assert.Empty(worksheet.Columns);
            worksheet.AddHiddenColumn(0);
            worksheet.AddHiddenColumn(1);
            worksheet.AddHiddenColumn(2);
            worksheet.SetAutoFilter("A1:C1");
            Assert.Equal(3, worksheet.Columns.Count);
            worksheet.SetColumnWidth("B", 66.6f);
            worksheet.ResetColumn(1); // Should not remove the column, since in a AutoFilter
            Assert.Equal(3, worksheet.Columns.Count);
            Assert.Contains(worksheet.Columns, item => item.Key == 1);
            Assert.False(worksheet.Columns[1].IsHidden);
            Assert.True(worksheet.Columns[1].HasAutoFilter);
            Assert.Equal(Worksheet.DEFAULT_COLUMN_WIDTH, worksheet.Columns[1].Width);
        }

        [Fact(DisplayName ="Test of the GetCurrentColumnNumber function")]
        public void GetCurrentColumnNumberTest()
        {
            Worksheet worksheet = new Worksheet();
            Assert.Equal(0, worksheet.GetCurrentColumnNumber());
            worksheet.CurrentCellDirection = Worksheet.CellDirection.ColumnToColumn;
            worksheet.AddNextCell("test");
            worksheet.AddNextCell("test");
            Assert.Equal(2, worksheet.GetCurrentColumnNumber());
            worksheet.CurrentCellDirection = Worksheet.CellDirection.RowToRow;
            worksheet.AddNextCell("test");
            worksheet.AddNextCell("test");
            Assert.Equal(2, worksheet.GetCurrentColumnNumber()); // should not change
            worksheet.GoToNextColumn();
            Assert.Equal(3, worksheet.GetCurrentColumnNumber());
            worksheet.GoToNextColumn(2);
            Assert.Equal(5, worksheet.GetCurrentColumnNumber());
            worksheet.GoToNextRow(2);
            Assert.Equal(0, worksheet.GetCurrentColumnNumber()); // should reset
        }

        [Theory(DisplayName = "Test of the GoToNextColumn function")]
        [InlineData(0, 0, 0)]
        [InlineData(0, 1, 1)]
        [InlineData(1, 1, 2)]
        [InlineData(3, 10, 13)]
        [InlineData(3, -1, 2)]
        [InlineData(3, -3, 0)]
        public void GoToNextColumnTest(int initialColumnNumber, int number, int expectedColumnNumber)
        {
            Worksheet worksheet = new Worksheet();
            worksheet.SetCurrentColumnNumber(initialColumnNumber);
            worksheet.GoToNextColumn(number);
            Assert.Equal(expectedColumnNumber, worksheet.GetCurrentColumnNumber());
        }

        [Theory(DisplayName = "Test of the GoToNextColumn function with the option to keep the row")]
        [InlineData("A1", 0, false, "A1")]
        [InlineData("A1", 0, true, "A1")]
        [InlineData("A1", 1, false, "B1")]
        [InlineData("A1", 1, true, "B1")]
        [InlineData("C10", 1, false, "D1")]
        [InlineData("C10", 1, true, "D10")]
        [InlineData("R5", 5, false, "W1")]
        [InlineData("R5", 5, true, "W5")]
        [InlineData("F5", -3, false, "C1")]
        [InlineData("F5", -3, true, "C5")]
        [InlineData("F5", -5, false, "A1")]
        [InlineData("F5", -5, true, "A5")]
        public void GoToNextColumnTest2(string initialAddress, int number, bool keepRowPosition, string expectedAddress)
        {
            Worksheet worksheet = new Worksheet();
            worksheet.SetCurrentCellAddress(initialAddress);
            worksheet.GoToNextColumn(number, keepRowPosition);
            Address expected = new Address(expectedAddress);
            Assert.Equal(expected.Column, worksheet.GetCurrentColumnNumber());
            Assert.Equal(expected.Row, worksheet.GetCurrentRowNumber());
        }

        [Theory(DisplayName = "Test of the failing GoToNextColumn function on invalid values")]
        [InlineData(0, -1)]
        [InlineData(10, -12)]
        [InlineData(0, 16384)]
        [InlineData(0, 20383)]
        public void GoToNextColumnFailTest(int initialValue, int value)
        {
            Worksheet worksheet = new Worksheet();
            worksheet.SetCurrentColumnNumber(initialValue);
            Assert.Equal(initialValue, worksheet.GetCurrentColumnNumber());
            Assert.Throws<RangeException>(() => worksheet.GoToNextColumn(value));
        }
        [Theory(DisplayName = "Test of the SetAutoFilter function on a start and end column")]
        [InlineData(0, 0, "A1:A1")]
        [InlineData(0, 5, "A1:F1")]
        [InlineData(1, 5, "B1:F1")]
        [InlineData(5, 1, "B1:F1")]
        public void SetAutoFilterTest(int startColumn, int endColumn, string expectedRange)
        {
            Worksheet worksheet = new Worksheet();
            Assert.Null(worksheet.AutoFilterRange);
            worksheet.SetAutoFilter(startColumn, endColumn);
            Assert.NotNull(worksheet.AutoFilterRange);
            Assert.Equal(expectedRange, worksheet.AutoFilterRange.ToString());
        }

        [Theory(DisplayName = "Test of the SetAutoFilter function on a range as string")]
        [InlineData("A1:A1", "A1:A1")]
        [InlineData("A1:F1", "A1:F1")]
        [InlineData("B1:F1", "B1:F1")]
        [InlineData("F1:B1", "B1:F1")]
        [InlineData("$B$1:$F$1", "B1:F1")]
        [InlineData("A1", "A1:A1")]
        public void SetAutoFilterTest2(string givenRange, string expectedRange)
        {
            Worksheet worksheet = new Worksheet();
            Assert.Null(worksheet.AutoFilterRange);
            worksheet.SetAutoFilter(givenRange);
            Assert.NotNull(worksheet.AutoFilterRange);
            Assert.Equal(expectedRange, worksheet.AutoFilterRange.ToString());
        }

        [Theory(DisplayName = "Test of the failing SetAutoFilter function on an invalid start and / or end column")]
        [InlineData(-1, 0)]
        [InlineData(0, -1)]
        [InlineData(-1, -1)]
        [InlineData(2, 16384)]
        [InlineData(16384, 2)]
        [InlineData(16384, 16384)]
        public void SetAutoFilterFailingTest(int startColumn, int endColumn)
        {
            Worksheet worksheet = new Worksheet();
            Assert.Throws<RangeException>(() => worksheet.SetAutoFilter(startColumn, endColumn));
        }

        [Theory(DisplayName = "Test of the failing SetAutoFilter function on an invalid string expression")]
        [InlineData("")]
        [InlineData(null)]
        [InlineData(":")]
        public void SetAutoFilterFailingTest2(string range)
        {
            Worksheet worksheet = new Worksheet();
            Assert.Throws<NanoXLSX.Shared.Exceptions.FormatException>(() => worksheet.SetAutoFilter(range));
        }

        [Theory(DisplayName = "Test of the SetColumnWidth function with column number and column address")]
        [InlineData(0f)]
        [InlineData(0.1f)]
        [InlineData(10f)]
        [InlineData(255f)]
        public void SetColumnWidthTest(float width)
        {
            Worksheet worksheet = new Worksheet();
            Assert.Empty(worksheet.Columns);
            worksheet.SetColumnWidth(0, width);
            worksheet.SetColumnWidth("B", width);
            Assert.Equal(2, worksheet.Columns.Count);
            Assert.Equal(width, worksheet.Columns[0].Width);
            Assert.Equal(width, worksheet.Columns[1].Width);
            worksheet.SetColumnWidth(0, Worksheet.DEFAULT_COLUMN_WIDTH);
            worksheet.SetColumnWidth("B", Worksheet.DEFAULT_COLUMN_WIDTH);
            Assert.Equal(2, worksheet.Columns.Count); // No removal so far
            Assert.Equal(Worksheet.DEFAULT_COLUMN_WIDTH, worksheet.Columns[0].Width);
            Assert.Equal(Worksheet.DEFAULT_COLUMN_WIDTH, worksheet.Columns[1].Width);
        }

        [Theory(DisplayName = "Test of the failing SetColumnWidth function with column number")]
        [InlineData(-1, 0f)]
        [InlineData(16384, 0.0f)]
        [InlineData(0, -10f)]
        [InlineData(0, 255.01f)]
        [InlineData(0, 500f)]
        public void SetColumnWidthFailTest(int columnNumber, float width)
        {
            Worksheet worksheet = new Worksheet();
            Assert.ThrowsAny<Exception>(() => worksheet.SetColumnWidth(columnNumber, width));
        }

        [Theory(DisplayName = "Test of the failing SetColumnWidth function with column address")]
        [InlineData(null, 0f)]
        [InlineData("", 0.0f)]
        [InlineData(":", 0.0f)]
        [InlineData("XFE", 0.0f)]
        [InlineData("A", -10f)]
        [InlineData("XFD", 255.01f)]
        [InlineData("A", 500f)]
        public void SetColumnWidthFailTest2(string columnAddress, float width)
        {
            Worksheet worksheet = new Worksheet();
            Assert.ThrowsAny<Exception>(() => worksheet.SetColumnWidth(columnAddress, width));
        }

        [Theory(DisplayName = "Test of the SetCurrentColumnNumber function")]
        [InlineData(0)]
        [InlineData(5)]
        [InlineData(16383)]
        public void SetCurrentColumnNumberTest(int column)
        {
            Worksheet worksheet = new Worksheet();
            Assert.Equal(0, worksheet.GetCurrentColumnNumber());
            worksheet.GoToNextColumn();
            worksheet.SetCurrentColumnNumber(column);
            Assert.Equal(column, worksheet.GetCurrentColumnNumber());
        }

        [Theory(DisplayName = "Test of the failing SetCurrentColumnNumber function")]
        [InlineData(-1)]
        [InlineData(-10)]
        [InlineData(16384)]
        public void SetCurrentColumnNumberFailTest(int column)
        {
            Worksheet worksheet = new Worksheet();
            Assert.Throws<RangeException>(() => worksheet.SetCurrentColumnNumber(column));
        }

        [Fact(DisplayName = "Test of the GetColumn function")]
        public void GetColumnTest()
        {
            Worksheet worksheet = new Worksheet();
            List<object> values = GetCellValues(worksheet);
            List<Cell> column = worksheet.GetColumn(1).ToList();
            AssertColumnValues(column, values);
        }

        [Fact(DisplayName = "Test of the GetColumn function with a column address")]
        public void GetColumnTest2()
        {
            Worksheet worksheet = new Worksheet();
            List<object> values = GetCellValues(worksheet);
            List<Cell> column = worksheet.GetColumn("B").ToList();
            AssertColumnValues(column, values);
        }

        [Fact(DisplayName = "Test of the GetColumn function when no values are applying")]
        public void GetColumnTest3()
        {
            Worksheet worksheet = new Worksheet();
            worksheet.AddCell(22, "A1");
            worksheet.AddCell(false, "C3");
            List<Cell> column = worksheet.GetColumn(1).ToList();
            Assert.Empty(column);
        }

        [Fact(DisplayName = "Test of the SetDefaultColumnStyle function with with a style and resetting it")]
        public void SetDefaultColumnStyleTest()
        {
            Worksheet worksheet = new Worksheet();
            Assert.Equal(0, worksheet.Columns.Count);
            Style style1 = BasicStyles.Font("Calibri Light", 13).Append(BasicStyles.BoldItalic);
            Style style2 = BasicStyles.Font("Arial", 11).Append(BasicStyles.DoubleUnderline);
            worksheet.SetColumnDefaultStyle(0, style1);
            worksheet.SetColumnDefaultStyle("B", style2);
            Assert.Equal(2, worksheet.Columns.Count);
            Assert.Equal(style1.GetHashCode(), worksheet.Columns[0].DefaultColumnStyle.GetHashCode());
            Assert.Equal(style2.GetHashCode(), worksheet.Columns[1].DefaultColumnStyle.GetHashCode());
            worksheet.SetColumnDefaultStyle(0, null);
            worksheet.SetColumnDefaultStyle("B", null);
            Assert.Equal(2, worksheet.Columns.Count); // No removal so far
            Assert.Null(worksheet.Columns[0].DefaultColumnStyle);
            Assert.Null(worksheet.Columns[1].DefaultColumnStyle);
        }
        [Theory(DisplayName = "Test of the failing SetDefaultColumnStyle function")]
        [InlineData(-1)]
        [InlineData(16384)]
        public void SetDefaultColumnStyleFailTest(int column)
        {
            Worksheet worksheet = new Worksheet();
            Assert.Throws<RangeException>(() => worksheet.SetColumnDefaultStyle(column, BasicStyles.Bold));
        }
        [Theory(DisplayName = "Test of the failing SetDefaultColumnStyle function with column address")]
        [InlineData(null)]
        [InlineData("")]
        [InlineData(":")]
        [InlineData("XFE")]
        public void SetDefaultColumnStyleFailTest2(string columnAddress)
        {
            Worksheet worksheet = new Worksheet();
            Assert.ThrowsAny<Exception>(() => worksheet.SetColumnDefaultStyle(columnAddress, BasicStyles.Bold));
        }

        private void AssertColumnValues(List<Cell> givenList, List<object> expectedValues)
        {
            Assert.Equal(expectedValues.Count, givenList.Count);
            for(int i = 0; i < expectedValues.Count; i++)
            {
                Assert.Equal(expectedValues[i], givenList[i].Value);
            }
        }

        private static List<object> GetCellValues(Worksheet worksheet)
        {
            List<object> expectedList = new List<object>();
            expectedList.Add(23);
            expectedList.Add("test");
            expectedList.Add(true);
            worksheet.AddCell(22, "A1");
            worksheet.AddCell(expectedList[0], "B1");
            worksheet.AddCell(expectedList[1], "B2");
            worksheet.AddCell(expectedList[2], "B3");
            worksheet.AddCell(false, "C2");
            return expectedList;
        }

    }
}
