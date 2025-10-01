using NanoXLSX;
using NanoXLSX.Styles;
using System;
using Xunit;

namespace NanoXLSX_Test.Worksheets
{
    public class AddCellFormulaTest
    {
        private Worksheet worksheet;

        [Fact(DisplayName = "Test of the AddCellFormula function with the only the value (with address and column/row invocation)")]
        public void AddCellFormulaTest1()
        {
            worksheet = WorksheetTest.InitWorksheet(worksheet, "D2", Worksheet.CellDirection.RowToRow);
            InvokeAddCellFormulaTest<int, int>("=B2", 2, 3, worksheet.AddCellFormula, "C4", 2, 4);
            Address address = new Address(3, 1);
            worksheet = WorksheetTest.InitWorksheet(worksheet, "R3", Worksheet.CellDirection.ColumnToColumn);
            InvokeAddCellFormulaTest<string>("=B2", address.GetAddress(), worksheet.AddCellFormula, "D2", 4, 1);
        }

        [Fact(DisplayName = "Test of the AddCellFormula function with value and Style (with address and column/row invocation)")]
        public void AddCellFormulaTest2()
        {
            worksheet = WorksheetTest.InitWorksheet(worksheet, "D2", Worksheet.CellDirection.RowToRow);
            InvokeAddCellFormulaTest<int, int, Style>("=B2", 2, 3, BasicStyles.BoldItalic, worksheet.AddCellFormula, "C4", 2, 4, BasicStyles.BoldItalic);
            Address address = new Address(3, 1);
            worksheet = WorksheetTest.InitWorksheet(worksheet, "R3", Worksheet.CellDirection.ColumnToColumn);
            InvokeAddCellFormulaTest<string, Style>("=B2", address.GetAddress(), BasicStyles.Bold, worksheet.AddCellFormula, "D2", 4, 1, BasicStyles.Bold);
        }

        [Fact(DisplayName = "Test of the AddCellFormula function with value and active worksheet style (with address and column/row invocation)")]
        public void AddCellFormulaTest3()
        {
            worksheet = WorksheetTest.InitWorksheet(worksheet, "D2", Worksheet.CellDirection.RowToRow, BasicStyles.BorderFrameHeader);
            InvokeAddCellFormulaTest<int, int>("=B2", 2, 3, worksheet.AddCellFormula, "C4", 2, 4, BasicStyles.BorderFrameHeader);
            Address address = new Address(3, 1);
            worksheet = WorksheetTest.InitWorksheet(worksheet, "R3", Worksheet.CellDirection.ColumnToColumn, BasicStyles.BorderFrameHeader);
            InvokeAddCellFormulaTest<string>("=B2", address.GetAddress(), worksheet.AddCellFormula, "D2", 4, 1, BasicStyles.BorderFrameHeader);
        }

        [Fact(DisplayName = "Test of the AddCell function for a nested cell object with a formula (with address and column/row invocation)")]
        public void AddCellFormulaTest4()
        {
            worksheet = WorksheetTest.InitWorksheet(worksheet, "D2", Worksheet.CellDirection.RowToRow);
            Cell cell = new Cell("=B2", Cell.CellType.FORMULA, "R1"); // Address should be replaced
            worksheet.AddCell(cell, 3, 1);
            WorksheetTest.AssertAddedCell(worksheet, 1, "D2", Cell.CellType.FORMULA, null, "=B2", 3, 2);
            worksheet = new Worksheet();
            worksheet = WorksheetTest.InitWorksheet(worksheet, "R3", Worksheet.CellDirection.ColumnToColumn);
            Address address = new Address(3, 1);
            worksheet.AddCell(cell, address.GetAddress());
            WorksheetTest.AssertAddedCell(worksheet, 1, "D2", Cell.CellType.FORMULA, null, "=B2", 4, 1);
        }

        [Fact(DisplayName = "Test of the AddCell function for a nested cell object with a formula and style (with address and column/row invocation)")]
        public void AddCellFormulaTest5()
        {
            worksheet = WorksheetTest.InitWorksheet(worksheet, "D2", Worksheet.CellDirection.RowToRow);
            Cell cell = new Cell("=B2", Cell.CellType.FORMULA, "R1"); // Address should be replaced
            worksheet.AddCell(cell, 3, 1, BasicStyles.Bold);
            WorksheetTest.AssertAddedCell(worksheet, 1, "D2", Cell.CellType.FORMULA, BasicStyles.Bold, "=B2", 3, 2);
            worksheet = new Worksheet();
            worksheet = WorksheetTest.InitWorksheet(worksheet, "R3", Worksheet.CellDirection.ColumnToColumn);
            Address address = new Address(3, 1);
            worksheet.AddCell(cell, address.GetAddress());
            WorksheetTest.AssertAddedCell(worksheet, 1, "D2", Cell.CellType.FORMULA, BasicStyles.Bold, "=B2", 4, 1);

            worksheet = WorksheetTest.InitWorksheet(worksheet, "D2", Worksheet.CellDirection.RowToRow);
            cell = new Cell("=B2", Cell.CellType.FORMULA, "R2");
            cell.SetStyle(BasicStyles.BorderFrame);
            Style mixedStyle = BasicStyles.BorderFrame;
            mixedStyle.Append(BasicStyles.Bold);
            worksheet.AddCell(cell, 3, 1, BasicStyles.Bold);
            WorksheetTest.AssertAddedCell(worksheet, 1, "D2", Cell.CellType.FORMULA, mixedStyle, "=B2", 3, 2);
            worksheet = new Worksheet();
            worksheet = WorksheetTest.InitWorksheet(worksheet, "R3", Worksheet.CellDirection.ColumnToColumn);
            worksheet.AddCell(cell, address.GetAddress(), BasicStyles.Bold);
            WorksheetTest.AssertAddedCell(worksheet, 1, "D2", Cell.CellType.FORMULA, mixedStyle, "=B2", 4, 1);
        }

        [Fact(DisplayName = "Test of the AddCell function for a nested cell object with a formula and active worksheet style (with address and column/row invocation)")]
        public void AddCellFormulaTest6()
        {
            worksheet = WorksheetTest.InitWorksheet(worksheet, "D2", Worksheet.CellDirection.RowToRow, BasicStyles.BorderFrame);
            Cell cell = new Cell("=B2", Cell.CellType.FORMULA, "R1"); // Address should be replaced
            worksheet.AddCell(cell, 3, 1);
            WorksheetTest.AssertAddedCell(worksheet, 1, "D2", Cell.CellType.FORMULA, BasicStyles.BorderFrame, "=B2", 3, 2);
            worksheet = WorksheetTest.InitWorksheet(worksheet, "D2", Worksheet.CellDirection.ColumnToColumn, BasicStyles.BorderFrame);
            Address address = new Address(3, 1);
            worksheet.AddCell(cell, address.GetAddress());
            WorksheetTest.AssertAddedCell(worksheet, 1, "D2", Cell.CellType.FORMULA, BasicStyles.BorderFrame, "=B2", 4, 1);

            worksheet = WorksheetTest.InitWorksheet(worksheet, "D2", Worksheet.CellDirection.RowToRow, BasicStyles.BorderFrame);
            cell = new Cell("=B2", Cell.CellType.FORMULA, "R2");
            cell.SetStyle(BasicStyles.Bold);
            Style mixedStyle = BasicStyles.BorderFrame;
            mixedStyle.Append(BasicStyles.Bold);
            worksheet.AddCell(cell, 3, 1);
            WorksheetTest.AssertAddedCell(worksheet, 1, "D2", Cell.CellType.FORMULA, mixedStyle, "=B2", 3, 2);
            worksheet = WorksheetTest.InitWorksheet(worksheet, "D2", Worksheet.CellDirection.ColumnToColumn, BasicStyles.BorderFrame);
            worksheet.AddCell(cell, address.GetAddress());
            WorksheetTest.AssertAddedCell(worksheet, 1, "D2", Cell.CellType.FORMULA, mixedStyle, "=B2", 4, 1);
        }

        [Theory(DisplayName = "Test of the AddCellFormula function with when changing the current cell direction (with address and column/row invocation)")]
        [InlineData("D2", 3, 1, Worksheet.CellDirection.RowToRow, 3, 2)]
        [InlineData("E7", 7, 2, Worksheet.CellDirection.ColumnToColumn, 8, 2)]
        [InlineData("C9", 10, 5, Worksheet.CellDirection.Disabled, 2, 8)]
        public void AddCellFormulaTest7(string worksheetAddress, int initialColumn, int initialRow, Worksheet.CellDirection cellDirection, int expectedNextColumn, int expectedNextRow)
        {
            Address initialAddress = new Address(initialColumn, initialRow);
            worksheet = WorksheetTest.InitWorksheet(worksheet, worksheetAddress, cellDirection);
            InvokeAddCellFormulaTest<int, int>("=B2", initialColumn, initialRow, worksheet.AddCellFormula, initialAddress.GetAddress(), expectedNextColumn, expectedNextRow);
            worksheet = WorksheetTest.InitWorksheet(worksheet, worksheetAddress, cellDirection);
            InvokeAddCellFormulaTest<string>("=B2", initialAddress.GetAddress(), worksheet.AddCellFormula, initialAddress.GetAddress(), expectedNextColumn, expectedNextRow);
        }

        private void InvokeAddCellFormulaTest<T1>(string value, T1 parameter1, Action<string, T1> action, string expectedAddress, int expectedNextColumn, int expectedNextRow, Style expectedStyle = null)
        {
            Assert.Empty(worksheet.Cells);
            action.Invoke(value, parameter1);
            AssertAddedFormulaCell(worksheet, 1, expectedAddress, expectedStyle, value, expectedNextColumn, expectedNextRow);
            worksheet = new Worksheet(); // Auto-reset
        }

        private void InvokeAddCellFormulaTest<T1, T2>(string value, T1 parameter1, T2 parameter2, Action<string, T1, T2> action, string expectedAddress, int expectedNextColumn, int expectedNextRow, Style expectedStyle = null)
        {
            Assert.Empty(worksheet.Cells);
            action.Invoke(value, parameter1, parameter2);
            AssertAddedFormulaCell(worksheet, 1, expectedAddress, expectedStyle, value, expectedNextColumn, expectedNextRow);
            worksheet = new Worksheet(); // Auto-reset
        }

        private void InvokeAddCellFormulaTest<T1, T2, T3>(string value, T1 parameter1, T2 parameter2, T3 parameter3, Action<string, T1, T2, T3> action, string expectedAddress, int expectedNextColumn, int expectedNextRow, Style expectedStyle = null)
        {
            Assert.Empty(worksheet.Cells);
            action.Invoke(value, parameter1, parameter2, parameter3);
            AssertAddedFormulaCell(worksheet, 1, expectedAddress, expectedStyle, value, expectedNextColumn, expectedNextRow);
            worksheet = new Worksheet(); // Auto-reset
        }

        private void AssertAddedFormulaCell(Worksheet worksheet, int numberOfEntries, string expectedAddress, Style expectedStyle, string expectedValue, int nextColumn, int nextRow)
        {
            Assert.Equal(numberOfEntries, worksheet.Cells.Count);
            Assert.Contains(worksheet.Cells, cell => cell.Key.Equals(expectedAddress));
            Assert.Equal(Cell.CellType.FORMULA, worksheet.Cells[expectedAddress].DataType);
            Assert.Equal(expectedValue, worksheet.Cells[expectedAddress].Value);
            if (expectedStyle == null)
            {
                Assert.Null(worksheet.Cells[expectedAddress].CellStyle);
            }
            else
            {
                Assert.True(expectedStyle.Equals(worksheet.Cells[expectedAddress].CellStyle));
            }
            Assert.Equal(nextColumn, worksheet.GetCurrentColumnNumber());
            Assert.Equal(nextRow, worksheet.GetCurrentRowNumber());
        }

    }
}
