using System;
using NanoXLSX.Styles;
using Xunit;

namespace NanoXLSX.Test.Core.WorksheetTest
{
    public class AddCellTest
    {
        private Worksheet worksheet;

        [Theory(DisplayName = "Test of the AddCell function with the only the value (with address and column/row invocation)")]
        [InlineData(null, 0, 0, Cell.CellType.EMPTY, "A1")]
        [InlineData("", 2, 2, Cell.CellType.STRING, "C3")]
        [InlineData("test", 5, 1, Cell.CellType.STRING, "F2")]
        [InlineData(17L, 16383, 0, Cell.CellType.NUMBER, "XFD1")]
        [InlineData(1.02d, 0, 1048575, Cell.CellType.NUMBER, "A1048576")]
        [InlineData(-22.3f, 16383, 1048575, Cell.CellType.NUMBER, "XFD1048576")]
        [InlineData(0, 0, 0, Cell.CellType.NUMBER, "A1")]
        [InlineData((byte)128, 2, 2, Cell.CellType.NUMBER, "C3")]
        [InlineData(true, 5, 1, Cell.CellType.BOOL, "F2")]
        [InlineData(false, 16383, 0, Cell.CellType.BOOL, "XFD1")]
        public void AddCellTest1(object value, int column, int row, Cell.CellType expectedType, string expectedAddress)
        {
            worksheet = WorksheetTest.InitWorksheet(worksheet, "D2", Worksheet.CellDirection.RowToRow);
            InvokeAddCellTest<int, int>(value, column, row, worksheet.AddCell, expectedType, expectedAddress, column, row + 1);
            Address address = new Address(column, row);
            worksheet = WorksheetTest.InitWorksheet(worksheet, "R3", Worksheet.CellDirection.ColumnToColumn);
            InvokeAddCellTest<string>(value, address.GetAddress(), worksheet.AddCell, expectedType, expectedAddress, column + 1, row);
        }

        [Theory(DisplayName = "Test of the AddCell function with value and Style (with address and column/row invocation)")]
        [InlineData(null, 0, 0, Cell.CellType.EMPTY, "A1")]
        [InlineData("", 2, 2, Cell.CellType.STRING, "C3")]
        [InlineData("test", 5, 1, Cell.CellType.STRING, "F2")]
        [InlineData(17L, 16383, 0, Cell.CellType.NUMBER, "XFD1")]
        [InlineData(1.02d, 0, 1048575, Cell.CellType.NUMBER, "A1048576")]
        [InlineData(-22.3f, 16383, 1048575, Cell.CellType.NUMBER, "XFD1048576")]
        [InlineData(0, 0, 0, Cell.CellType.NUMBER, "A1")]
        [InlineData((byte)128, 2, 2, Cell.CellType.NUMBER, "C3")]
        [InlineData(true, 5, 1, Cell.CellType.BOOL, "F2")]
        [InlineData(false, 16383, 0, Cell.CellType.BOOL, "XFD1")]
        public void AddCellTest2(object value, int column, int row, Cell.CellType expectedType, string expectedAddress)
        {
            worksheet = WorksheetTest.InitWorksheet(worksheet, "D2", Worksheet.CellDirection.RowToRow);
            InvokeAddCellTest<int, int, Style>(value, column, row, BasicStyles.BoldItalic, worksheet.AddCell, expectedType, expectedAddress, column, row + 1, BasicStyles.BoldItalic);
            Address address = new Address(column, row);
            worksheet = WorksheetTest.InitWorksheet(worksheet, "R3", Worksheet.CellDirection.ColumnToColumn);
            InvokeAddCellTest<string, Style>(value, address.GetAddress(), BasicStyles.Bold, worksheet.AddCell, expectedType, expectedAddress, column + 1, row, BasicStyles.Bold);
        }

        [Fact(DisplayName = "Test of the AddCell function for DateTime and TimeSpan (with address and column/row invocation)")]
        public void AddCellTest3()
        {
            worksheet = WorksheetTest.InitWorksheet(worksheet, "D2", Worksheet.CellDirection.RowToRow);
            DateTime date = new DateTime(2020, 6, 10, 11, 12, 22);
            InvokeAddCellTest<int, int>(date, 5, 1, worksheet.AddCell, Cell.CellType.DATE, "F2", 5, 2, BasicStyles.DateFormat);
            Address address = new Address(5, 1);
            worksheet = WorksheetTest.InitWorksheet(worksheet, "R3", Worksheet.CellDirection.ColumnToColumn);
            InvokeAddCellTest<string>(date, address.GetAddress(), worksheet.AddCell, Cell.CellType.DATE, "F2", 6, 1, BasicStyles.DateFormat);

            worksheet = WorksheetTest.InitWorksheet(worksheet, "S9", Worksheet.CellDirection.RowToRow);
            TimeSpan time = new TimeSpan(6, 22, 13);
            InvokeAddCellTest<int, int>(time, 5, 1, worksheet.AddCell, Cell.CellType.TIME, "F2", 5, 2, BasicStyles.TimeFormat);
            worksheet = WorksheetTest.InitWorksheet(worksheet, "V6", Worksheet.CellDirection.ColumnToColumn);
            InvokeAddCellTest<string>(time, address.GetAddress(), worksheet.AddCell, Cell.CellType.TIME, "F2", 6, 1, BasicStyles.TimeFormat);
        }

        [Fact(DisplayName = "Test of the AddCell function for DateTime and TimeSpan with styles (with address and column/row invocation)")]
        public void AddCellTest4()
        {
            worksheet = WorksheetTest.InitWorksheet(worksheet, "D2", Worksheet.CellDirection.RowToRow);
            DateTime date = new DateTime(2020, 6, 10, 11, 12, 22);
            Style mixedStyle = BasicStyles.DateFormat;
            mixedStyle.Append(BasicStyles.Bold);
            InvokeAddCellTest<int, int, Style>(date, 5, 1, BasicStyles.Bold, worksheet.AddCell, Cell.CellType.DATE, "F2", 5, 2, mixedStyle);
            Address address = new Address(5, 1);
            worksheet = WorksheetTest.InitWorksheet(worksheet, "R3", Worksheet.CellDirection.ColumnToColumn);
            InvokeAddCellTest<string, Style>(date, address.GetAddress(), BasicStyles.Bold, worksheet.AddCell, Cell.CellType.DATE, "F2", 6, 1, mixedStyle);

            worksheet = WorksheetTest.InitWorksheet(worksheet, "S9", Worksheet.CellDirection.RowToRow);
            TimeSpan time = new TimeSpan(6, 22, 13);
            mixedStyle = BasicStyles.TimeFormat;
            mixedStyle.Append(BasicStyles.Underline);
            InvokeAddCellTest<int, int, Style>(time, 5, 1, BasicStyles.Underline, worksheet.AddCell, Cell.CellType.TIME, "F2", 5, 2, mixedStyle);
            worksheet = WorksheetTest.InitWorksheet(worksheet, "V6", Worksheet.CellDirection.ColumnToColumn);
            InvokeAddCellTest<string, Style>(time, address.GetAddress(), BasicStyles.Underline, worksheet.AddCell, Cell.CellType.TIME, "F2", 6, 1, mixedStyle);
        }


        [Theory(DisplayName = "Test of the AddCell function with value and active worksheet style (with address and column/row invocation)")]
        [InlineData(null, 0, 0, Cell.CellType.EMPTY, "A1")]
        [InlineData("", 2, 2, Cell.CellType.STRING, "C3")]
        [InlineData("test", 5, 1, Cell.CellType.STRING, "F2")]
        [InlineData(17L, 16383, 0, Cell.CellType.NUMBER, "XFD1")]
        [InlineData(1.02d, 0, 1048575, Cell.CellType.NUMBER, "A1048576")]
        [InlineData(-22.3f, 16383, 1048575, Cell.CellType.NUMBER, "XFD1048576")]
        [InlineData(0, 0, 0, Cell.CellType.NUMBER, "A1")]
        [InlineData((byte)128, 2, 2, Cell.CellType.NUMBER, "C3")]
        [InlineData(true, 5, 1, Cell.CellType.BOOL, "F2")]
        [InlineData(false, 16383, 0, Cell.CellType.BOOL, "XFD1")]
        public void AddCellTest5(object value, int column, int row, Cell.CellType expectedType, string expectedAddress)
        {
            worksheet = WorksheetTest.InitWorksheet(worksheet, "D2", Worksheet.CellDirection.RowToRow, BasicStyles.BorderFrameHeader);
            InvokeAddCellTest<int, int>(value, column, row, worksheet.AddCell, expectedType, expectedAddress, column, row + 1, BasicStyles.BorderFrameHeader);
            Address address = new Address(column, row);
            worksheet = WorksheetTest.InitWorksheet(worksheet, "R3", Worksheet.CellDirection.ColumnToColumn, BasicStyles.BorderFrameHeader);
            InvokeAddCellTest<string>(value, address.GetAddress(), worksheet.AddCell, expectedType, expectedAddress, column + 1, row, BasicStyles.BorderFrameHeader);
        }

        [Fact(DisplayName = "Test of the AddCell function for DateTime and TimeSpan with active worksheet style (with address and column/row invocation)")]
        public void AddCellTest6()
        {

            worksheet = WorksheetTest.InitWorksheet(worksheet, "D2", Worksheet.CellDirection.RowToRow, BasicStyles.BorderFrameHeader);
            DateTime date = new DateTime(2020, 6, 10, 11, 12, 22);
            Style mixedStyle = BasicStyles.DateFormat;
            mixedStyle.Append(BasicStyles.BorderFrameHeader);
            InvokeAddCellTest<int, int>(date, 5, 1, worksheet.AddCell, Cell.CellType.DATE, "F2", 5, 2, mixedStyle);
            Address address = new Address(5, 1);
            worksheet = WorksheetTest.InitWorksheet(worksheet, "R3", Worksheet.CellDirection.ColumnToColumn, BasicStyles.BorderFrameHeader);
            InvokeAddCellTest<string>(date, address.GetAddress(), worksheet.AddCell, Cell.CellType.DATE, "F2", 6, 1, mixedStyle);

            worksheet = WorksheetTest.InitWorksheet(worksheet, "S9", Worksheet.CellDirection.RowToRow, BasicStyles.Underline);
            TimeSpan time = new TimeSpan(6, 22, 13);
            mixedStyle = BasicStyles.TimeFormat;
            mixedStyle.Append(BasicStyles.Underline);
            InvokeAddCellTest<int, int>(time, 5, 1, worksheet.AddCell, Cell.CellType.TIME, "F2", 5, 2, mixedStyle);
            worksheet = WorksheetTest.InitWorksheet(worksheet, "V6", Worksheet.CellDirection.ColumnToColumn, BasicStyles.Underline);
            InvokeAddCellTest<string>(time, address.GetAddress(), worksheet.AddCell, Cell.CellType.TIME, "F2", 6, 1, mixedStyle);
        }


        [Fact(DisplayName = "Test of the AddCell function for a nested cell object (with address and column/row invocation)")]
        public void AddCellTest7()
        {
            worksheet = WorksheetTest.InitWorksheet(worksheet, "D2", Worksheet.CellDirection.RowToRow);
            Cell cell = new Cell(33.3d, Cell.CellType.NUMBER, "R1"); // Address should be replaced
            worksheet.AddCell(cell, 3, 1);
            WorksheetTest.AssertAddedCell(worksheet, 1, "D2", Cell.CellType.NUMBER, null, 33.3d, 3, 2);
            worksheet = new Worksheet();
            worksheet = WorksheetTest.InitWorksheet(worksheet, "R3", Worksheet.CellDirection.ColumnToColumn);
            Address address = new Address(3, 1);
            worksheet.AddCell(cell, address.GetAddress());
            WorksheetTest.AssertAddedCell(worksheet, 1, "D2", Cell.CellType.NUMBER, null, 33.3d, 4, 1);
        }


        [Fact(DisplayName = "Test of the AddCell function for a nested cell object and style (with address and column/row invocation)")]
        public void AddCellTest8()
        {
            worksheet = WorksheetTest.InitWorksheet(worksheet, "D2", Worksheet.CellDirection.RowToRow);
            Cell cell = new Cell(33.3d, Cell.CellType.NUMBER, "R1"); // Address should be replaced
            worksheet.AddCell(cell, 3, 1, BasicStyles.Bold);
            WorksheetTest.AssertAddedCell(worksheet, 1, "D2", Cell.CellType.NUMBER, BasicStyles.Bold, 33.3d, 3, 2);
            worksheet = new Worksheet();
            worksheet = WorksheetTest.InitWorksheet(worksheet, "R3", Worksheet.CellDirection.ColumnToColumn);
            Address address = new Address(3, 1);
            worksheet.AddCell(cell, address.GetAddress());
            WorksheetTest.AssertAddedCell(worksheet, 1, "D2", Cell.CellType.NUMBER, BasicStyles.Bold, 33.3d, 4, 1);

            worksheet = WorksheetTest.InitWorksheet(worksheet, "D2", Worksheet.CellDirection.RowToRow);
            cell = new Cell("test", Cell.CellType.STRING, "R2");
            cell.SetStyle(BasicStyles.BorderFrame);
            Style mixedStyle = BasicStyles.BorderFrame;
            mixedStyle.Append(BasicStyles.Bold);
            worksheet.AddCell(cell, 3, 1, BasicStyles.Bold);
            WorksheetTest.AssertAddedCell(worksheet, 1, "D2", Cell.CellType.STRING, mixedStyle, "test", 3, 2);
            worksheet = new Worksheet();
            worksheet = WorksheetTest.InitWorksheet(worksheet, "R3", Worksheet.CellDirection.ColumnToColumn);
            worksheet.AddCell(cell, address.GetAddress(), BasicStyles.Bold);
            WorksheetTest.AssertAddedCell(worksheet, 1, "D2", Cell.CellType.STRING, mixedStyle, "test", 4, 1);
        }


        [Fact(DisplayName = "Test of the AddCell function for a nested cell object and active worksheet style (with address and column/row invocation)")]
        public void AddCellTest9()
        {
            worksheet = WorksheetTest.InitWorksheet(worksheet, "D2", Worksheet.CellDirection.RowToRow, BasicStyles.BorderFrame);
            Cell cell = new Cell(33.3d, Cell.CellType.NUMBER, "R1"); // Address should be replaced
            worksheet.AddCell(cell, 3, 1);
            WorksheetTest.AssertAddedCell(worksheet, 1, "D2", Cell.CellType.NUMBER, BasicStyles.BorderFrame, 33.3d, 3, 2);
            worksheet = WorksheetTest.InitWorksheet(worksheet, "D2", Worksheet.CellDirection.ColumnToColumn, BasicStyles.BorderFrame);
            Address address = new Address(3, 1);
            worksheet.AddCell(cell, address.GetAddress());
            WorksheetTest.AssertAddedCell(worksheet, 1, "D2", Cell.CellType.NUMBER, BasicStyles.BorderFrame, 33.3d, 4, 1);

            worksheet = WorksheetTest.InitWorksheet(worksheet, "D2", Worksheet.CellDirection.RowToRow, BasicStyles.BorderFrame);
            cell = new Cell("test", Cell.CellType.STRING, "R2");
            cell.SetStyle(BasicStyles.Bold);
            Style mixedStyle = BasicStyles.BorderFrame;
            mixedStyle.Append(BasicStyles.Bold);
            worksheet.AddCell(cell, 3, 1);
            WorksheetTest.AssertAddedCell(worksheet, 1, "D2", Cell.CellType.STRING, mixedStyle, "test", 3, 2);
            worksheet = WorksheetTest.InitWorksheet(worksheet, "D2", Worksheet.CellDirection.ColumnToColumn, BasicStyles.BorderFrame);
            worksheet.AddCell(cell, address.GetAddress());
            WorksheetTest.AssertAddedCell(worksheet, 1, "D2", Cell.CellType.STRING, mixedStyle, "test", 4, 1);
        }


        [Theory(DisplayName = "Test of the AddCell function with when changing the current cell direction (with address and column/row invocation)")]
        [InlineData("D2", 3, 1, Worksheet.CellDirection.RowToRow, 3, 2)]
        [InlineData("E7", 7, 2, Worksheet.CellDirection.ColumnToColumn, 8, 2)]
        [InlineData("C9", 10, 5, Worksheet.CellDirection.Disabled, 2, 8)]
        public void AddCellTest10(string worksheetAddress, int initialColumn, int initialRow, Worksheet.CellDirection cellDirection, int expectedNextColumn, int expectedNextRow)
        {
            Address initialAddress = new Address(initialColumn, initialRow);
            worksheet = WorksheetTest.InitWorksheet(worksheet, worksheetAddress, cellDirection);
            InvokeAddCellTest<int, int>("test", initialColumn, initialRow, worksheet.AddCell, Cell.CellType.STRING, initialAddress.GetAddress(), expectedNextColumn, expectedNextRow);
            worksheet = WorksheetTest.InitWorksheet(worksheet, worksheetAddress, cellDirection);
            InvokeAddCellTest<string>("test", initialAddress.GetAddress(), worksheet.AddCell, Cell.CellType.STRING, initialAddress.GetAddress(), expectedNextColumn, expectedNextRow);
        }

        [Fact(DisplayName = "Test of the AddCell function where an existing cell is overwritten")]
        public void AddCellOverwriteTest()
        {
            Worksheet worksheet2 = new Worksheet();
            worksheet2.AddCell("test", "C2");
            Assert.Equal(Cell.CellType.STRING, worksheet2.Cells["C2"].DataType);
            Assert.Equal("test", worksheet2.Cells["C2"].Value);
            worksheet2.AddCell(22, "C2");
            Assert.Equal(Cell.CellType.NUMBER, worksheet2.Cells["C2"].DataType);
            Assert.Equal(22, worksheet2.Cells["C2"].Value);
            Assert.Single(worksheet2.Cells);
        }

        [Fact(DisplayName = "Test of the AddCell function where existing cells are overwritten and the old cells where dates and times")]
        public void AddCellOverwriteTest2()
        {
            Worksheet worksheet2 = new Worksheet();
            DateTime date = new DateTime(2020, 10, 5, 4, 11, 12);
            TimeSpan time = new TimeSpan(11, 12, 13);
            worksheet2.AddCell(date, "C2");
            worksheet2.AddCell(time, "C3");
            Assert.Equal(Cell.CellType.DATE, worksheet2.Cells["C2"].DataType);
            Assert.Equal(date, worksheet2.Cells["C2"].Value);
            Assert.True(BasicStyles.DateFormat.Equals(worksheet2.Cells["C2"].CellStyle));
            Assert.Equal(Cell.CellType.TIME, worksheet2.Cells["C3"].DataType);
            Assert.Equal(time, worksheet2.Cells["C3"].Value);
            Assert.True(BasicStyles.TimeFormat.Equals(worksheet2.Cells["C3"].CellStyle));
            worksheet2.AddCell(22, "C2");
            worksheet2.AddCell("test", "C3");
            Assert.Equal(Cell.CellType.NUMBER, worksheet2.Cells["C2"].DataType);
            Assert.Equal(22, worksheet2.Cells["C2"].Value);
            Assert.Null(worksheet2.Cells["C2"].CellStyle);
            Assert.Equal(Cell.CellType.STRING, worksheet2.Cells["C3"].DataType);
            Assert.Equal("test", worksheet2.Cells["C3"].Value);
            Assert.Null(worksheet2.Cells["C3"].CellStyle);
            Assert.Equal(2, worksheet2.Cells.Count);
        }


        private void InvokeAddCellTest<T1>(object value, T1 parameter1, Action<object, T1> action, Cell.CellType expectedType, string expectedAddress, int expectedNextColumn, int expectedNextRow, Style expectedStyle = null)
        {
            Assert.Empty(worksheet.Cells);
            action.Invoke(value, parameter1);
            WorksheetTest.AssertAddedCell(worksheet, 1, expectedAddress, expectedType, expectedStyle, value, expectedNextColumn, expectedNextRow);
            worksheet = new Worksheet(); // Auto-reset
        }

        private void InvokeAddCellTest<T1, T2>(object value, T1 parameter1, T2 parameter2, Action<object, T1, T2> action, Cell.CellType expectedType, string expectedAddress, int expectedNextColumn, int expectedNextRow, Style expectedStyle = null)
        {
            Assert.Empty(worksheet.Cells);
            action.Invoke(value, parameter1, parameter2);
            WorksheetTest.AssertAddedCell(worksheet, 1, expectedAddress, expectedType, expectedStyle, value, expectedNextColumn, expectedNextRow);
            worksheet = new Worksheet(); // Auto-reset
        }

        private void InvokeAddCellTest<T1, T2, T3>(object value, T1 parameter1, T2 parameter2, T3 parameter3, Action<object, T1, T2, T3> action, Cell.CellType expectedType, string expectedAddress, int expectedNextColumn, int expectedNextRow, Style expectedStyle = null)
        {
            Assert.Empty(worksheet.Cells);
            action.Invoke(value, parameter1, parameter2, parameter3);
            WorksheetTest.AssertAddedCell(worksheet, 1, expectedAddress, expectedType, expectedStyle, value, expectedNextColumn, expectedNextRow);
            worksheet = new Worksheet(); // Auto-reset
        }

    }
}
