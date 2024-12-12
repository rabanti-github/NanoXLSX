using NanoXLSX.Styles;
using Xunit;

namespace NanoXLSX.Test.Worksheets
{

    public class AddNextCellFormulaTest
    {
        private Worksheet worksheet;

        [Fact(DisplayName = "Test of the AddNextCellFormula function with only the value")]
        public void AddNextCellFormulaTest1()
        {
            worksheet = WorksheetTest.InitWorksheet(worksheet, "D2", Worksheet.CellDirection.RowToRow);
            Assert.Empty(worksheet.Cells);
            worksheet.AddNextCellFormula("=B2");
            WorksheetTest.AssertAddedCell(worksheet, 1, "D2", Cell.CellType.FORMULA, null, "=B2", 3, 2);
            worksheet = WorksheetTest.InitWorksheet(worksheet, "E3", Worksheet.CellDirection.ColumnToColumn);
            worksheet.AddNextCellFormula("=B2");
            WorksheetTest.AssertAddedCell(worksheet, 2, "E3", Cell.CellType.FORMULA, null, "=B2", 5, 2);
        }

        [Fact(DisplayName = "Test of the AddNextCellFormula function with value and Style")]
        public void AddNextCellFormulaTest2()
        {
            worksheet = WorksheetTest.InitWorksheet(worksheet, "D2", Worksheet.CellDirection.RowToRow);
            Assert.Empty(worksheet.Cells);
            worksheet.AddNextCellFormula("=B2", BasicStyles.BoldItalic);
            WorksheetTest.AssertAddedCell(worksheet, 1, "D2", Cell.CellType.FORMULA, BasicStyles.BoldItalic, "=B2", 3, 2);
            worksheet = WorksheetTest.InitWorksheet(worksheet, "E3", Worksheet.CellDirection.ColumnToColumn);
            worksheet.AddNextCellFormula("=B2", BasicStyles.Bold);
            WorksheetTest.AssertAddedCell(worksheet, 2, "E3", Cell.CellType.FORMULA, BasicStyles.Bold, "=B2", 5, 2);
        }

        [Fact(DisplayName = "Test of the AddNextCellFormula function with value and active worksheet style")]
        public void AddNextCellFormulaTest3()
        {
            worksheet = WorksheetTest.InitWorksheet(worksheet, "D2", Worksheet.CellDirection.RowToRow, BasicStyles.BorderFrameHeader);
            Assert.Empty(worksheet.Cells);
            worksheet.AddNextCellFormula("=B2");
            WorksheetTest.AssertAddedCell(worksheet, 1, "D2", Cell.CellType.FORMULA, BasicStyles.BorderFrameHeader, "=B2", 3, 2);
        }

        [Fact(DisplayName = "Test of the AddNextCell function for a nested cell object, if the cell is a formula")]
        public void AddNextCellFormulaTest5()
        {
            Cell cell = new Cell("=B2", Cell.CellType.FORMULA, "R1"); // Address should be replaced
            worksheet = WorksheetTest.InitWorksheet(worksheet, "D2", Worksheet.CellDirection.RowToRow);
            worksheet.AddNextCell(cell);
            WorksheetTest.AssertAddedCell(worksheet, 1, "D2", Cell.CellType.FORMULA, null, "=B2", 3, 2);
        }

        [Fact(DisplayName = "Test of the AddNextCell function for a nested cell object and style, if the cell is a formula")]
        public void AddNextCellFormulaTest6()
        {
            Cell cell = new Cell("=B2", Cell.CellType.FORMULA, "R1"); // Address should be replaced
            worksheet = WorksheetTest.InitWorksheet(worksheet, "D2", Worksheet.CellDirection.RowToRow);
            worksheet.AddNextCell(cell, BasicStyles.Bold);
            WorksheetTest.AssertAddedCell(worksheet, 1, "D2", Cell.CellType.FORMULA, BasicStyles.Bold, "=B2", 3, 2);
            cell = new Cell("=B2", Cell.CellType.FORMULA, "R2");
            cell.SetStyle(BasicStyles.BorderFrame);
            Style mixedStyle = BasicStyles.BorderFrame;
            mixedStyle.Append(BasicStyles.Bold);
            worksheet.AddNextCell(cell, BasicStyles.Bold);
            WorksheetTest.AssertAddedCell(worksheet, 2, "D3", Cell.CellType.FORMULA, mixedStyle, "=B2", 3, 3);
        }

        [Fact(DisplayName = "Test of the AddNextCell function for a nested cell object and active worksheet style, if the cell is a formula")]
        public void AddNextCellFormulaTest7()
        {
            worksheet = WorksheetTest.InitWorksheet(worksheet, "D2", Worksheet.CellDirection.RowToRow, BasicStyles.Bold);
            Cell cell = new Cell("=B2", Cell.CellType.FORMULA, "R1"); // Address should be replaced
            worksheet.AddNextCell(cell);
            WorksheetTest.AssertAddedCell(worksheet, 1, "D2", Cell.CellType.FORMULA, BasicStyles.Bold, "=B2", 3, 2);
            cell = new Cell("=B2", Cell.CellType.FORMULA, "R2");
            cell.SetStyle(BasicStyles.BorderFrame);
            Style mixedStyle = BasicStyles.BorderFrame;
            mixedStyle.Append(BasicStyles.Bold);
            worksheet.AddNextCell(cell);
            WorksheetTest.AssertAddedCell(worksheet, 2, "D3", Cell.CellType.FORMULA, mixedStyle, "=B2", 3, 3);
        }

        [Fact(DisplayName = "Test of the AddNextCellFormula function with when changing the current cell direction")]
        public void AddNextCellFormulaTest8()
        {
            Worksheet worksheet = new Worksheet();
            worksheet = WorksheetTest.InitWorksheet(worksheet, "D2", Worksheet.CellDirection.RowToRow);
            worksheet.AddNextCellFormula("=B2");
            WorksheetTest.AssertAddedCell(worksheet, 1, "D2", Cell.CellType.FORMULA, null, "=B2", 3, 2);
            worksheet = WorksheetTest.InitWorksheet(worksheet, "E3", Worksheet.CellDirection.ColumnToColumn);
            worksheet.AddNextCellFormula("=B2");
            WorksheetTest.AssertAddedCell(worksheet, 2, "E3", Cell.CellType.FORMULA, null, "=B2", 5, 2);
            worksheet = WorksheetTest.InitWorksheet(worksheet, "F5", Worksheet.CellDirection.Disabled);
            worksheet.AddNextCellFormula("=B2");
            WorksheetTest.AssertAddedCell(worksheet, 3, "F5", Cell.CellType.FORMULA, null, "=B2", 5, 4);
        }
    }
}
