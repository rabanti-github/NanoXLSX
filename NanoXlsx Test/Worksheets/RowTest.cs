using NanoXLSX;
using NanoXLSX.Shared.Exceptions;
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

        [Fact(DisplayName = "Test of the GetCurrentColumnNumber function")]
        public void GetCurrentRowNumberTest()
        {
            Worksheet worksheet = new Worksheet();
            Assert.Equal(0, worksheet.GetCurrentRowNumber());
            worksheet.CurrentCellDirection = Worksheet.CellDirection.RowToRow;
            worksheet.AddNextCell("test");
            worksheet.AddNextCell("test");
            Assert.Equal(2, worksheet.GetCurrentRowNumber());
            worksheet.CurrentCellDirection = Worksheet.CellDirection.ColumnToColumn;
            worksheet.AddNextCell("test");
            worksheet.AddNextCell("test");
            Assert.Equal(2, worksheet.GetCurrentRowNumber()); // should not change
            worksheet.GoToNextRow();
            Assert.Equal(3, worksheet.GetCurrentRowNumber());
            worksheet.GoToNextRow(2);
            Assert.Equal(5, worksheet.GetCurrentRowNumber());
            worksheet.GoToNextColumn(2);
            Assert.Equal(0, worksheet.GetCurrentRowNumber()); // should reset
        }

        [Theory(DisplayName = "Test of the GoToNextRow function")]
        [InlineData(0, 0, 0)]
        [InlineData(0, 1, 1)]
        [InlineData(1, 1, 2)]
        [InlineData(3, 10, 13)]
        [InlineData(3, -1, 2)]
        [InlineData(3, -3, 0)]
        public void GoToNextRowTest(int initialRowNumber, int number, int expectedRowNumber)
        {
            Worksheet worksheet = new Worksheet();
            worksheet.SetCurrentRowNumber(initialRowNumber);
            worksheet.GoToNextRow(number);
            Assert.Equal(expectedRowNumber, worksheet.GetCurrentRowNumber());
        }

        [Theory(DisplayName = "Test of the GoToNextRow function with the option to keep the column")]
        [InlineData("A1", 0, false, "A1")]
        [InlineData("A1", 0, true, "A1")]
        [InlineData("A1", 1, false, "A2")]
        [InlineData("A1", 1, true, "A2")]
        [InlineData("C10", 1, false, "A11")]
        [InlineData("C10", 1, true, "C11")]
        [InlineData("R5", 5, false, "A10")]
        [InlineData("R5", 5, true, "R10")]
        [InlineData("F5", -3, false, "A2")]
        [InlineData("F5", -3, true, "F2")]
        [InlineData("F5", -4, false, "A1")]
        [InlineData("F5", -4, true, "F1")]
        public void GoToNextRowTest2(string initialAddress, int number, bool keepColumnPosition, string expectedAddress)
        {
            Worksheet worksheet = new Worksheet();
            worksheet.SetCurrentCellAddress(initialAddress);
            worksheet.GoToNextRow(number, keepColumnPosition);
            Address expected = new Address(expectedAddress);
            Assert.Equal(expected.Column, worksheet.GetCurrentColumnNumber());
            Assert.Equal(expected.Row, worksheet.GetCurrentRowNumber());
        }

        [Theory(DisplayName = "Test of the failing GoToNextRow function on invalid values")]
        [InlineData(0, -1)]
        [InlineData(10, -12)]
        [InlineData(0, 1048576)]
        [InlineData(0, 1248575)]
        public void GoToNextRowFailTest(int initialValue, int value)
        {
            Worksheet worksheet = new Worksheet();
            worksheet.SetCurrentRowNumber(initialValue);
            Assert.Equal(initialValue, worksheet.GetCurrentRowNumber());
            Assert.Throws<RangeException>(() => worksheet.GoToNextRow(value));
        }

        [Fact(DisplayName = "Test of the RemoveRowHeight function")]
        public void RemoveRowHeightTest()
        {
            Worksheet worksheet = new Worksheet();
            worksheet.SetRowHeight(2, 22.2f);
            worksheet.SetRowHeight(4, 33.3f);
            Assert.Equal(2, worksheet.RowHeights.Count);
            worksheet.RemoveRowHeight(2);
            Assert.Single(worksheet.RowHeights);
            worksheet.RemoveRowHeight(3); // Should not cause anything
            worksheet.RemoveRowHeight(-1); // Should not cause anything
            Assert.Single(worksheet.RowHeights);
        }

        [Theory(DisplayName = "Test of the SetCurrentRowNumber function")]
        [InlineData(0)]
        [InlineData(3)]
        [InlineData(1048575)]
        public void SetCurrentRowNumberTest(int row)
        {
            Worksheet worksheet = new Worksheet();
            Assert.Equal(0, worksheet.GetCurrentRowNumber());
            worksheet.GoToNextRow();
            worksheet.SetCurrentRowNumber(row);
            Assert.Equal(row, worksheet.GetCurrentRowNumber());
        }

        [Theory(DisplayName = "Test of the failing SetCurrentRowNumber function")]
        [InlineData(-1)]
        [InlineData(-10)]
        [InlineData(1048576)]
        public void SetCurrentRowNumberFailTest(int row)
        {
            Worksheet worksheet = new Worksheet();
            Assert.Throws<RangeException>(() => worksheet.SetCurrentRowNumber(row));
        }

        [Theory(DisplayName = "Test of the SetRowHeight function")]
        [InlineData(0f)]
        [InlineData(0.1f)]
        [InlineData(10f)]
        [InlineData(255f)]
        public void SetRowHeightTest(float height)
        {
            Worksheet worksheet = new Worksheet();
            Assert.Empty(worksheet.RowHeights);
            worksheet.SetRowHeight(0, height);
            Assert.Single(worksheet.RowHeights);
            Assert.Equal(height, worksheet.RowHeights[0]);
            worksheet.SetRowHeight(0, Worksheet.DEFAULT_ROW_HEIGHT);
            Assert.Single(worksheet.RowHeights); // No removal so far
            Assert.Equal(Worksheet.DEFAULT_ROW_HEIGHT, worksheet.RowHeights[0]);
        }

        [Theory(DisplayName = "Test of the failing SetRowHeight function")]
        [InlineData(-1, 0f)]
        [InlineData(1048576, 0.0f)]
        [InlineData(0, -10f)]
        [InlineData(0, 409.51f)]
        [InlineData(0, 500f)]
        public void SetRowHeightFailTest(int rowNumber, float height)
        {
            Worksheet worksheet = new Worksheet();
            Assert.Throws<RangeException>(() => worksheet.SetRowHeight(rowNumber, height));
        }

        [Fact(DisplayName ="Test of the GetRow function")]
        public void GetRowTest()
        {
            Worksheet worksheet = new Worksheet();
            worksheet.AddCell(22, "B1");
            worksheet.AddCell(23, "B2");
            worksheet.AddCell("test", "C2");
            worksheet.AddCell(true, "D2");
            worksheet.AddCell(false, "B3");
            List<Cell> row = worksheet.GetRow(1).ToList();
            Assert.Equal(3, row.Count());
            Assert.Equal(23, row[0].Value);
            Assert.Equal("test", row[1].Value);
            Assert.Equal(true, row[2].Value);
        }

        [Fact(DisplayName = "Test of the GetRow function when no values are applying")]
        public void GetRowTest2()
        {
            Worksheet worksheet = new Worksheet();
            worksheet.AddCell(22, "B1");
            worksheet.AddCell(false, "B3");
            List<Cell> row = worksheet.GetRow(1).ToList();
            Assert.Empty(row);
        }

    }
}
