using NanoXLSX;
using NanoXLSX.Styles;
using System;
using Xunit;

namespace NanoXLSX_Test.Worksheets
{

    public class AddNextCellTest
    {
        private Worksheet worksheet;

        [Theory(DisplayName = "Test of the AddNextCell function with only the value")]
        [InlineData(null, Cell.CellType.EMPTY)]
        [InlineData("", Cell.CellType.STRING)]
        [InlineData("test", Cell.CellType.STRING)]
        [InlineData(17L, Cell.CellType.NUMBER)]
        [InlineData(1.02d, Cell.CellType.NUMBER)]
        [InlineData(-22.3f, Cell.CellType.NUMBER)]
        [InlineData(0, Cell.CellType.NUMBER)]
        [InlineData((byte)128, Cell.CellType.NUMBER)]
        [InlineData(true, Cell.CellType.BOOL)]
        [InlineData(false, Cell.CellType.BOOL)]
        public void AddNextCellTest1(object value, Cell.CellType expectedType)
        {
            worksheet = WorksheetTest.InitWorksheet(worksheet, "D2", Worksheet.CellDirection.RowToRow);
            Assert.Empty(worksheet.Cells);
            worksheet.AddNextCell(value);
            WorksheetTest.AssertAddedCell(worksheet, 1, "D2", expectedType, null, value, 3, 2);
            worksheet = WorksheetTest.InitWorksheet(worksheet, "E3", Worksheet.CellDirection.ColumnToColumn);
            worksheet.AddNextCell(value);
            WorksheetTest.AssertAddedCell(worksheet, 2, "E3", expectedType, null, value, 5, 2);
        }

        [Theory(DisplayName = "Test of the AddNextCell function with value and Style")]
        [InlineData(null, Cell.CellType.EMPTY)]
        [InlineData("", Cell.CellType.STRING)]
        [InlineData("test", Cell.CellType.STRING)]
        [InlineData(17L, Cell.CellType.NUMBER)]
        [InlineData(1.02d, Cell.CellType.NUMBER)]
        [InlineData(-22.3f, Cell.CellType.NUMBER)]
        [InlineData(0, Cell.CellType.NUMBER)]
        [InlineData((byte)128, Cell.CellType.NUMBER)]
        [InlineData(true, Cell.CellType.BOOL)]
        [InlineData(false, Cell.CellType.BOOL)]
        public void AddNextCellTest2(object value, Cell.CellType expectedType)
        {
            worksheet = WorksheetTest.InitWorksheet(worksheet, "D2", Worksheet.CellDirection.RowToRow);
            Assert.Empty(worksheet.Cells);
            worksheet.AddNextCell(value, BasicStyles.BoldItalic);
            WorksheetTest.AssertAddedCell(worksheet, 1, "D2", expectedType, BasicStyles.BoldItalic, value, 3, 2);
            worksheet = WorksheetTest.InitWorksheet(worksheet, "E3", Worksheet.CellDirection.ColumnToColumn);
            worksheet.AddNextCell(value, BasicStyles.Bold);
            WorksheetTest.AssertAddedCell(worksheet, 2, "E3", expectedType, BasicStyles.Bold, value, 5, 2);
        }

        [Fact(DisplayName = "Test of the AddNextCell function for DateTime and TimeSpan")]
        public void AddNextCellTest3()
        {
            worksheet = WorksheetTest.InitWorksheet(worksheet, "D2", Worksheet.CellDirection.RowToRow);
            Assert.Empty(worksheet.Cells);
            DateTime date = new DateTime(2020, 6, 10, 11, 12, 22);
            worksheet.AddNextCell(date);
            WorksheetTest.AssertAddedCell(worksheet, 1, "D2", Cell.CellType.DATE, BasicStyles.DateFormat, date, 3, 2);
            worksheet = WorksheetTest.InitWorksheet(worksheet, "E3", Worksheet.CellDirection.ColumnToColumn);
            TimeSpan time = new TimeSpan(6, 22, 13);
            worksheet.AddNextCell(time);
            WorksheetTest.AssertAddedCell(worksheet, 2, "E3", Cell.CellType.TIME, BasicStyles.TimeFormat, time, 5, 2);
        }

        [Fact(DisplayName = "Test of the AddNextCell function for DateTime and TimeSpan with styles")]
        public void AddNextCellTest4()
        {
            worksheet = WorksheetTest.InitWorksheet(worksheet, "D2", Worksheet.CellDirection.RowToRow);
            Assert.Empty(worksheet.Cells);
            DateTime date = new DateTime(2020, 6, 10, 11, 12, 22);
            worksheet.AddNextCell(date, BasicStyles.Bold);
            Style mixedStyle = BasicStyles.DateFormat;
            mixedStyle.Append(BasicStyles.Bold);
            WorksheetTest.AssertAddedCell(worksheet, 1, "D2", Cell.CellType.DATE, mixedStyle, date, 3, 2);
            worksheet = WorksheetTest.InitWorksheet(worksheet, "E3", Worksheet.CellDirection.ColumnToColumn);
            TimeSpan time = new TimeSpan(6, 22, 13);
            worksheet.AddNextCell(time, BasicStyles.Underline);
            mixedStyle = BasicStyles.TimeFormat;
            mixedStyle.Append(BasicStyles.Underline);
            WorksheetTest.AssertAddedCell(worksheet, 2, "E3", Cell.CellType.TIME, mixedStyle, time, 5, 2);
        }

        [Theory(DisplayName = "Test of the AddNextCell function with value and active worksheet style")]
        [InlineData(null, Cell.CellType.EMPTY)]
        [InlineData("", Cell.CellType.STRING)]
        [InlineData("test", Cell.CellType.STRING)]
        [InlineData(17L, Cell.CellType.NUMBER)]
        [InlineData(1.02d, Cell.CellType.NUMBER)]
        [InlineData(-22.3f, Cell.CellType.NUMBER)]
        [InlineData(0, Cell.CellType.NUMBER)]
        [InlineData((byte)128, Cell.CellType.NUMBER)]
        [InlineData(true, Cell.CellType.BOOL)]
        [InlineData(false, Cell.CellType.BOOL)]
        public void AddNextCellTest5(object value, Cell.CellType expectedType)
        {
            worksheet = WorksheetTest.InitWorksheet(worksheet, "D2", Worksheet.CellDirection.RowToRow, BasicStyles.BorderFrameHeader);
            Assert.Empty(worksheet.Cells);
            worksheet.AddNextCell(value);
            WorksheetTest.AssertAddedCell(worksheet, 1, "D2", expectedType, BasicStyles.BorderFrameHeader, value, 3, 2);
        }

        [Fact(DisplayName = "Test of the AddNextCell function for DateTime and TimeSpan with active worksheet style")]
        public void AddNextCellTest6()
        {
            worksheet = WorksheetTest.InitWorksheet(worksheet, "D2", Worksheet.CellDirection.RowToRow, BasicStyles.BorderFrameHeader);
            Assert.Empty(worksheet.Cells);
            DateTime date = new DateTime(2020, 6, 10, 11, 12, 22);
            worksheet.AddNextCell(date);
            Style mixedStyle = BasicStyles.DateFormat;
            mixedStyle.Append(BasicStyles.BorderFrameHeader);
            WorksheetTest.AssertAddedCell(worksheet, 1, "D2", Cell.CellType.DATE, mixedStyle, date, 3, 2);
            worksheet = WorksheetTest.InitWorksheet(worksheet, "E3", Worksheet.CellDirection.ColumnToColumn);
            TimeSpan time = new TimeSpan(6, 22, 13);
            worksheet.AddNextCell(time);
            mixedStyle = BasicStyles.TimeFormat;
            mixedStyle.Append(BasicStyles.BorderFrameHeader);
            WorksheetTest.AssertAddedCell(worksheet, 2, "E3", Cell.CellType.TIME, mixedStyle, time, 5, 2);
        }

        [Fact(DisplayName = "Test of the AddNextCell function for a nested cell object")]
        public void AddNextCellTest7()
        {
            Cell cell = new Cell(33.3d, Cell.CellType.NUMBER, "R1"); // Address should be replaced
            worksheet = WorksheetTest.InitWorksheet(worksheet, "D2", Worksheet.CellDirection.RowToRow);
            worksheet.AddNextCell(cell);
            WorksheetTest.AssertAddedCell(worksheet, 1, "D2", Cell.CellType.NUMBER, null, 33.3d, 3, 2);
        }

        [Fact(DisplayName = "Test of the AddNextCell function for a nested cell object and style")]
        public void AddNextCellTest8()
        {
            Cell cell = new Cell(33.3d, Cell.CellType.NUMBER, "R1"); // Address should be replaced
            worksheet = WorksheetTest.InitWorksheet(worksheet, "D2", Worksheet.CellDirection.RowToRow);
            worksheet.AddNextCell(cell, BasicStyles.Bold);
            WorksheetTest.AssertAddedCell(worksheet, 1, "D2", Cell.CellType.NUMBER, BasicStyles.Bold, 33.3d, 3, 2);
            cell = new Cell("test", Cell.CellType.STRING, "R2");
            cell.SetStyle(BasicStyles.BorderFrame);
            Style mixedStyle = BasicStyles.BorderFrame;
            mixedStyle.Append(BasicStyles.Bold);
            worksheet.AddNextCell(cell, BasicStyles.Bold);
            WorksheetTest.AssertAddedCell(worksheet, 2, "D3", Cell.CellType.STRING, mixedStyle, "test", 3, 3);
        }

        [Fact(DisplayName = "Test of the AddNextCell function for a nested cell object and active worksheet style")]
        public void AddNextCellTest9()
        {
            worksheet = WorksheetTest.InitWorksheet(worksheet, "D2", Worksheet.CellDirection.RowToRow, BasicStyles.Bold);
            Cell cell = new Cell(33.3d, Cell.CellType.NUMBER, "R1"); // Address should be replaced
            worksheet.AddNextCell(cell);
            WorksheetTest.AssertAddedCell(worksheet, 1, "D2", Cell.CellType.NUMBER, BasicStyles.Bold, 33.3d, 3, 2);
            cell = new Cell("test", Cell.CellType.STRING, "R2");
            cell.SetStyle(BasicStyles.BorderFrame);
            Style mixedStyle = BasicStyles.BorderFrame;
            mixedStyle.Append(BasicStyles.Bold);
            worksheet.AddNextCell(cell);
            WorksheetTest.AssertAddedCell(worksheet, 2, "D3", Cell.CellType.STRING, mixedStyle, "test", 3, 3);
        }

        [Fact(DisplayName = "Test of the AddNextCell function with when changing the current cell direction")]
        public void AddNextCellTest10()
        {
            Worksheet worksheet = new Worksheet();
            worksheet = WorksheetTest.InitWorksheet(worksheet, "D2", Worksheet.CellDirection.RowToRow);
            worksheet.AddNextCell("test");
            WorksheetTest.AssertAddedCell(worksheet, 1, "D2", Cell.CellType.STRING, null, "test", 3, 2);
            worksheet = WorksheetTest.InitWorksheet(worksheet, "E3", Worksheet.CellDirection.ColumnToColumn);
            worksheet.AddNextCell("test");
            WorksheetTest.AssertAddedCell(worksheet, 2, "E3", Cell.CellType.STRING, null, "test", 5, 2);
            worksheet = WorksheetTest.InitWorksheet(worksheet, "F5", Worksheet.CellDirection.Disabled);
            worksheet.AddNextCell("test");
            WorksheetTest.AssertAddedCell(worksheet, 3, "F5", Cell.CellType.STRING, null, "test", 5, 4);
        }
    }
}
