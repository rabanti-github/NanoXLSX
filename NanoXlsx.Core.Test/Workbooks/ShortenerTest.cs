using System;
using System.Collections.Generic;
using NanoXLSX.Shared.Exceptions;
using NanoXLSX.Styles;
using Xunit;

namespace NanoXLSX.Test
{
    public class ShortenerTest
    {
        [Fact(DisplayName = "Test of the SetCurrentWorksheet function")]
        public void SetCurrentWorksheetTest()
        {
            Workbook workbook = new Workbook("Sheet1");
            Worksheet worksheet = new Worksheet("Sheet2");
            workbook.AddWorksheet(worksheet);
            workbook.SetCurrentWorksheet("Sheet1");
            Assert.Equal("Sheet1", workbook.CurrentWorksheet.SheetName);
            workbook.WS.SetCurrentWorksheet(worksheet);
            Assert.Equal("Sheet2", workbook.CurrentWorksheet.SheetName);
        }

        [Fact(DisplayName = "Test of the failing SetCurrentWorksheet function on a unreferenced worksheet")]
        public void SetCurrentWorksheetFailTest()
        {
            Workbook workbook = new Workbook("Sheet1");
            Worksheet worksheet = new Worksheet("Sheet2");
            Assert.Throws<WorksheetException>(() => workbook.WS.SetCurrentWorksheet(worksheet));
        }

        [Fact(DisplayName = "Test of the failing SetCurrentWorksheet function on a null object")]
        public void SetCurrentWorksheetFailTest2()
        {
            Workbook workbook = new Workbook("Sheet1");
            Assert.Throws<WorksheetException>(() => workbook.WS.SetCurrentWorksheet(null));
        }


        [Fact(DisplayName = "Test of the Value function")]
        public void ValueTest()
        {
            Workbook workbook = new Workbook("Sheet1");
            Dictionary<string, object> values = new Dictionary<string, object>();
            values.Add("A1", "Test");
            values.Add("B1", 22);
            AssertValue(workbook, workbook.WS.Value, null, values, Worksheet.CellDirection.ColumnToColumn, 0, 0, 2, 0, null);

            values.Clear();
            values.Add("C3", "Test2");
            values.Add("C4", 22.2);
            AssertValue(workbook, workbook.WS.Value, null, values, Worksheet.CellDirection.RowToRow, 2, 2, 2, 4, null);
        }

        [Fact(DisplayName = "Test of the Value function with a style")]
        public void ValueTest2()
        {
            Workbook workbook = new Workbook("Sheet1");
            Dictionary<string, object> values = new Dictionary<string, object>();
            values.Add("A1", true);
            values.Add("B1", "");
            Style style = BasicStyles.BoldItalic;
            AssertValue(workbook, null, workbook.WS.Value, values, Worksheet.CellDirection.ColumnToColumn, 0, 0, 2, 0, style);

            values.Clear();
            values.Add("C3", -22.3);
            values.Add("C4", false);
            style = BasicStyles.DoubleUnderline;
            AssertValue(workbook, null, workbook.WS.Value, values, Worksheet.CellDirection.RowToRow, 2, 2, 2, 4, style);
        }

        [Fact(DisplayName = "Test of the Formula function")]
        public void FormulaTest()
        {
            Workbook workbook = new Workbook("Sheet1");
            Dictionary<string, string> values = new Dictionary<string, string>();
            values.Add("A1", "=A3");
            values.Add("B1", "=ROUNDDOWN(22.1)");
            AssertValue<string>(workbook, workbook.WS.Formula, null, values, Worksheet.CellDirection.ColumnToColumn, 0, 0, 2, 0, null);

            values.Clear();
            values.Add("C3", "=C3");
            values.Add("C4", "=ROUNDDOWN(11.1)");
            AssertValue(workbook, workbook.WS.Value, null, values, Worksheet.CellDirection.RowToRow, 2, 2, 2, 4, null);
        }

        [Fact(DisplayName = "Test of the Formula function with a style")]
        public void FormulaTest2()
        {
            Workbook workbook = new Workbook("Sheet1");
            Dictionary<string, string> values = new Dictionary<string, string>();
            values.Add("A1", "=A3");
            values.Add("B1", "=ROUNDDOWN(22.1)");
            Style style = BasicStyles.BoldItalic;
            AssertValue(workbook, null, workbook.WS.Formula, values, Worksheet.CellDirection.ColumnToColumn, 0, 0, 2, 0, style);

            values.Clear();
            values.Add("C3", "=C3");
            values.Add("C4", "=ROUNDDOWN(11.1)");
            style = BasicStyles.DoubleUnderline;
            AssertValue(workbook, null, workbook.WS.Formula, values, Worksheet.CellDirection.RowToRow, 2, 2, 2, 4, style);
        }

        [Fact(DisplayName = "Test of the Down function")]
        public void DownTest()
        {
            Workbook workbook = new Workbook("Sheet1");
            Assert.Equal(0, workbook.CurrentWorksheet.GetCurrentColumnNumber());
            Assert.Equal(0, workbook.CurrentWorksheet.GetCurrentRowNumber());
            workbook.WS.Down();
            Assert.Equal(0, workbook.CurrentWorksheet.GetCurrentColumnNumber());
            Assert.Equal(1, workbook.CurrentWorksheet.GetCurrentRowNumber());
        }

        [Theory(DisplayName = "Test of the Down function with a row number")]
        [InlineData(0, 0, 0, 0, 0)]
        [InlineData(0, 0, 1, 0, 1)]
        [InlineData(5, 5, 5, 0, 10)]
        [InlineData(5, 5, -2, 0, 3)]
        [InlineData(5, 5, -5, 0, 0)]
        public void DownTest2(int startColumn, int startRow, int number, int expectedColumn, int expectedRow)
        {
            Workbook workbook = new Workbook("Sheet1");
            AssertJumpTo(workbook, workbook.WS.Down, startColumn, startRow, number, expectedColumn, expectedRow);
        }

        [Theory(DisplayName = "Test of the Down function with a row number and the option to keep the column position")]
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
        public void DownTest3(string initialAddress, int number, bool keepColumn, string expectedAdrress)
        {
            Workbook workbook = new Workbook("Sheet1");
            AssertJumpKeep(workbook, workbook.WS.Down, initialAddress, number, keepColumn, expectedAdrress);
        }

        [Fact(DisplayName = "Test of the failing Down function with a negative row number")]
        public void DownFailingTest()
        {
            Workbook workbook = new Workbook("Sheet1");
            Assert.Equal(0, workbook.CurrentWorksheet.GetCurrentColumnNumber());
            Assert.Equal(0, workbook.CurrentWorksheet.GetCurrentRowNumber());
            Assert.Throws<RangeException>(() => workbook.WS.Down(-2));
        }

        [Fact(DisplayName = "Test of the Up function")]
        public void UpTest()
        {
            Workbook workbook = new Workbook("Sheet1");
            workbook.CurrentWorksheet.SetCurrentCellAddress("C4");
            Assert.Equal(2, workbook.CurrentWorksheet.GetCurrentColumnNumber());
            Assert.Equal(3, workbook.CurrentWorksheet.GetCurrentRowNumber());
            workbook.WS.Up();
            Assert.Equal(0, workbook.CurrentWorksheet.GetCurrentColumnNumber());
            Assert.Equal(2, workbook.CurrentWorksheet.GetCurrentRowNumber());
        }

        [Theory(DisplayName = "Test of the Up function with a row number")]
        [InlineData(0, 0, 0, 0, 0)]
        [InlineData(1, 1, 1, 0, 0)]
        [InlineData(5, 10, 5, 0, 5)]
        [InlineData(5, 5, -2, 0, 7)]
        [InlineData(5, 5, 5, 0, 0)]
        public void UpTest2(int startColumn, int startRow, int number, int expectedColumn, int expectedRow)
        {
            Workbook workbook = new Workbook("Sheet1");
            AssertJumpTo(workbook, workbook.WS.Up, startColumn, startRow, number, expectedColumn, expectedRow);
        }

        [Theory(DisplayName = "Test of the Up function with a row number and the option to keep the column position")]
        [InlineData("A1", 0, false, "A1")]
        [InlineData("A1", 0, true, "A1")]
        [InlineData("A2", 1, false, "A1")]
        [InlineData("A2", 1, true, "A1")]
        [InlineData("C10", 1, false, "A9")]
        [InlineData("C10", 1, true, "C9")]
        [InlineData("R10", 5, false, "A5")]
        [InlineData("R10", 5, true, "R5")]
        [InlineData("F5", -3, false, "A8")]
        [InlineData("F5", -3, true, "F8")]
        [InlineData("F5", 4, false, "A1")]
        [InlineData("F5", 4, true, "F1")]
        public void UpTest3(string initialAddress, int number, bool keepColumn, string expectedAdrress)
        {
            Workbook workbook = new Workbook("Sheet1");
            AssertJumpKeep(workbook, workbook.WS.Up, initialAddress, number, keepColumn, expectedAdrress);
        }

        [Fact(DisplayName = "Test of the failing Up function with a negative row number")]
        public void UpFailingTest()
        {
            Workbook workbook = new Workbook("Sheet1");
            Assert.Equal(0, workbook.CurrentWorksheet.GetCurrentColumnNumber());
            Assert.Equal(0, workbook.CurrentWorksheet.GetCurrentRowNumber());
            Assert.Throws<RangeException>(() => workbook.WS.Up(2));
        }

        [Fact(DisplayName = "Test of the Right function")]
        public void RightTest()
        {
            Workbook workbook = new Workbook("Sheet1");
            Assert.Equal(0, workbook.CurrentWorksheet.GetCurrentColumnNumber());
            Assert.Equal(0, workbook.CurrentWorksheet.GetCurrentRowNumber());
            workbook.WS.Right();
            Assert.Equal(1, workbook.CurrentWorksheet.GetCurrentColumnNumber());
            Assert.Equal(0, workbook.CurrentWorksheet.GetCurrentRowNumber());
        }

        [Theory(DisplayName = "Test of the Right function with a column number")]
        [InlineData(0, 0, 0, 0, 0)]
        [InlineData(0, 0, 1, 1, 0)]
        [InlineData(5, 5, 5, 10, 0)]
        [InlineData(5, 5, -2, 3, 0)]
        [InlineData(5, 5, -5, 0, 0)]
        public void RightTest2(int startColumn, int startRow, int number, int expectedColumn, int expectedRow)
        {
            Workbook workbook = new Workbook("Sheet1");
            AssertJumpTo(workbook, workbook.WS.Right, startColumn, startRow, number, expectedColumn, expectedRow);
        }

        [Theory(DisplayName = "Test of the Right function with a column number and the option to keep the row position")]
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
        public void RightTest3(string initialAddress, int number, bool keepColumn, string expectedAdrress)
        {
            Workbook workbook = new Workbook("Sheet1");
            AssertJumpKeep(workbook, workbook.WS.Right, initialAddress, number, keepColumn, expectedAdrress);
        }

        [Fact(DisplayName = "Test of the failing Right function with a negative row number")]
        public void RightFailingTest()
        {
            Workbook workbook = new Workbook("Sheet1");
            Assert.Equal(0, workbook.CurrentWorksheet.GetCurrentColumnNumber());
            Assert.Equal(0, workbook.CurrentWorksheet.GetCurrentRowNumber());
            Assert.Throws<RangeException>(() => workbook.WS.Right(-2));
        }

        [Fact(DisplayName = "Test of the Left function")]
        public void LeftTest()
        {
            Workbook workbook = new Workbook("Sheet1");
            workbook.CurrentWorksheet.SetCurrentCellAddress("D4");
            Assert.Equal(3, workbook.CurrentWorksheet.GetCurrentColumnNumber());
            Assert.Equal(3, workbook.CurrentWorksheet.GetCurrentRowNumber());
            workbook.WS.Left();
            Assert.Equal(2, workbook.CurrentWorksheet.GetCurrentColumnNumber());
            Assert.Equal(0, workbook.CurrentWorksheet.GetCurrentRowNumber());
        }

        [Theory(DisplayName = "Test of the Left function with a column number")]
        [InlineData(0, 0, 0, 0, 0)]
        [InlineData(1, 1, 1, 0, 0)]
        [InlineData(5, 5, 2, 3, 0)]
        [InlineData(5, 5, -2, 7, 0)]
        [InlineData(5, 5, 5, 0, 0)]
        public void LeftTest2(int startColumn, int startRow, int number, int expectedColumn, int expectedRow)
        {
            Workbook workbook = new Workbook("Sheet1");
            AssertJumpTo(workbook, workbook.WS.Left, startColumn, startRow, number, expectedColumn, expectedRow);
        }

        [Theory(DisplayName = "Test of the Left function with a column number and the option to keep the row position")]
        [InlineData("A1", 0, false, "A1")]
        [InlineData("A1", 0, true, "A1")]
        [InlineData("B1", 1, false, "A1")]
        [InlineData("B1", 1, true, "A1")]
        [InlineData("C10", 1, false, "B1")]
        [InlineData("C10", 1, true, "B10")]
        [InlineData("R5", 5, false, "M1")]
        [InlineData("R5", 5, true, "M5")]
        [InlineData("F5", -3, false, "I1")]
        [InlineData("F5", -3, true, "I5")]
        [InlineData("F5", 5, false, "A1")]
        [InlineData("F5", 5, true, "A5")]
        public void LeftTest3(string initialAddress, int number, bool keepColumn, string expectedAdrress)
        {
            Workbook workbook = new Workbook("Sheet1");
            AssertJumpKeep(workbook, workbook.WS.Left, initialAddress, number, keepColumn, expectedAdrress);
        }

        [Fact(DisplayName = "Test of the failing Left function with a negative row number")]
        public void LeftFailingTest()
        {
            Workbook workbook = new Workbook("Sheet1");
            Assert.Equal(0, workbook.CurrentWorksheet.GetCurrentColumnNumber());
            Assert.Equal(0, workbook.CurrentWorksheet.GetCurrentRowNumber());
            Assert.Throws<RangeException>(() => workbook.WS.Left(2));
        }

        // For code coverage
        [Fact(DisplayName = "Singular Test of the NullCheck method")]
        public void NullCheckTest()
        {
            Workbook workbook = new Workbook(); // No worksheet created
            Assert.Throws<WorksheetException>(() => workbook.WS.Value(22));
        }

        private void AssertJumpTo(Workbook workbook, JumpDelegate action, int startColumn, int startRow, int number, int expectedColumn, int expectedRow)
        {
            workbook.CurrentWorksheet.SetCurrentColumnNumber(startColumn);
            workbook.CurrentWorksheet.SetCurrentRowNumber(startRow);
            Assert.Equal(startColumn, workbook.CurrentWorksheet.GetCurrentColumnNumber());
            Assert.Equal(startRow, workbook.CurrentWorksheet.GetCurrentRowNumber());
            action.Invoke(number);
            Assert.Equal(expectedColumn, workbook.CurrentWorksheet.GetCurrentColumnNumber());
            Assert.Equal(expectedRow, workbook.CurrentWorksheet.GetCurrentRowNumber());
        }

        private void AssertJumpKeep(Workbook workbook, JumpKeepDelegate action, string initialAddress, int number, bool keepOther, string expectedAddress)
        {
            Address initial = new Address(initialAddress);
            workbook.CurrentWorksheet.SetCurrentCellAddress(initial.Column, initial.Row);

            Assert.Equal(initial.Column, workbook.CurrentWorksheet.GetCurrentColumnNumber());
            Assert.Equal(initial.Row, workbook.CurrentWorksheet.GetCurrentRowNumber());
            action.Invoke(number, keepOther);
            Address expected = new Address(expectedAddress);
            Assert.Equal(expected.Column, workbook.CurrentWorksheet.GetCurrentColumnNumber());
            Assert.Equal(expected.Row, workbook.CurrentWorksheet.GetCurrentRowNumber());
        }

        private void AssertValue<T>(Workbook workbook, Action<T> action, Action<T, Style> styleAction, Dictionary<string, T> values, Worksheet.CellDirection direction, int startColumn, int startRow, int expectedEndColumn, int expectedEndRow, Style style)
        {
            workbook.CurrentWorksheet.SetCurrentColumnNumber(startColumn);
            workbook.CurrentWorksheet.SetCurrentRowNumber(startRow);
            workbook.CurrentWorksheet.CurrentCellDirection = direction;

            foreach (KeyValuePair<string, T> cell in values)
            {
                if (style == null)
                {
                    action.Invoke(cell.Value);
                }
                else
                {
                    styleAction.Invoke(cell.Value, style);
                }
            }

            foreach (KeyValuePair<string, T> cell in values)
            {
                Address address = new Address(cell.Key);
                T value = (T)workbook.CurrentWorksheet.GetCell(address).Value;
                Assert.Equal(cell.Value, value);
                if (style != null)
                {
                    Assert.Equal(style.GetHashCode(), workbook.CurrentWorksheet.GetCell(address).CellStyle.GetHashCode());
                }
            }
            Assert.Equal(expectedEndColumn, workbook.CurrentWorksheet.GetCurrentColumnNumber());
            Assert.Equal(expectedEndRow, workbook.CurrentWorksheet.GetCurrentRowNumber());
        }

        delegate void JumpDelegate(int number, bool keepOther = false);

        delegate void JumpKeepDelegate(int number, bool keepOther);

    }
}
