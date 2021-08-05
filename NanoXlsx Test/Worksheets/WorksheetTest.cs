using NanoXLSX;
using NanoXLSX.Styles;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xunit;

namespace NanoXLSX_Test.Worksheets
{
    // Ensure that these tests are executed sequentially, since static repository methods may be called 
    [Collection(nameof(SequentialCollection))]
    public class WorksheetTest
    {

        [Fact(DisplayName = "Test of the default constructor")]
        public void ConstructorTest()
        {
            Worksheet worksheet = new Worksheet();
            AssertConstructorBasics(worksheet);
            Assert.Null(worksheet.WorkbookReference);
            Assert.Equal(0, worksheet.SheetID);
        }

        [Theory(DisplayName = "Test of the constructor with parameters")]
        [InlineData(".", 1)]
        [InlineData(" ", 2)]
        [InlineData("Test", 10)]
        [InlineData("xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx", 255)]
        public void ConstructorTest2(string name, int id)
        {
            Workbook workbook = new Workbook("test.xlsx", "sheet2");
            Worksheet worksheet = new Worksheet(name, id, workbook);
            AssertConstructorBasics(worksheet);
            Assert.NotNull(worksheet.WorkbookReference);
            Assert.Equal("test.xlsx", worksheet.WorkbookReference.Filename);
            Assert.Equal(id, worksheet.SheetID);
        }

        [Theory(DisplayName = "Test failing of the constructor if provided with invalid values")]
        [InlineData("", 1)]
        [InlineData(null, 1)]
        [InlineData("[", 1)]
        [InlineData("................................", 0)]
        [InlineData("Test", 0)]
        [InlineData("Test", -1)]
        public void ConstructorFailingTest(string name, int id)
        {
            Workbook workbook = new Workbook("test.xlsx", "sheet2");
            Assert.Throws<NanoXLSX.Exceptions.FormatException>(() => new Worksheet(name, id, workbook));
        }

        [Fact(DisplayName = "Test of the get function of the AutoFilterRang property")]
        public void AutoFilterRangTest()
        {
            Worksheet worksheet = new Worksheet();
            Assert.Null(worksheet.AutoFilterRange);
            worksheet.SetAutoFilter("B2:D4");
            NanoXLSX.Range ExpectedRange = new NanoXLSX.Range("B1:D1"); // Function reduces range to row 1
            Assert.Equal(ExpectedRange, worksheet.AutoFilterRange);
            worksheet.RemoveAutoFilter();
            Assert.Null(worksheet.AutoFilterRange);
        }

        [Fact(DisplayName = "Test of the get function of the Cells property")]
        public void CellsTest()
        {
            Worksheet worksheet = new Worksheet();
            Assert.NotNull(worksheet.Cells);
            Assert.Empty(worksheet.Cells);
            worksheet.AddCell("test", "C3");
            worksheet.AddCell(22, "D4");
            Assert.Equal(2, worksheet.Cells.Count);
            Assert.Contains(worksheet.Cells, item => (item.Key.Equals("C3") && item.Value.Value.Equals("test")));
            Assert.Contains(worksheet.Cells, item => (item.Key.Equals("D4") && item.Value.Value.Equals(22)));
            worksheet.RemoveCell("C3");
            Assert.Single(worksheet.Cells);
            Assert.Contains(worksheet.Cells, item => (item.Key.Equals("D4") && item.Value.Value.Equals(22)));
        }

        [Fact(DisplayName = "Test of the get function of the Columns property")]
        public void ColumnsTest()
        {
            Worksheet worksheet = new Worksheet();
            Assert.NotNull(worksheet.Columns);
            Assert.Empty(worksheet.Columns);
            worksheet.SetColumnWidth("B", 11f);
            worksheet.SetColumnWidth("C", 0.7f);
            Assert.Equal(2, worksheet.Columns.Count);
            Assert.Contains(worksheet.Columns, item => (item.Key.Equals(1) && item.Value.Width.Equals(11f)));
            Assert.Contains(worksheet.Columns, item => (item.Key.Equals(2) && item.Value.Width.Equals(0.7f)));
            worksheet.ResetColumn(1);
            Assert.Single(worksheet.Columns);
            Assert.Contains(worksheet.Columns, item => (item.Key.Equals(2) && item.Value.Width.Equals(0.7f)));
        }

        [Theory(DisplayName = "Test of the CurrentCellDirection property")]
        [InlineData(Worksheet.CellDirection.ColumnToColumn, 2, 7, 3, 7)]
        [InlineData(Worksheet.CellDirection.RowToRow, 2, 7, 2, 8)]
        [InlineData(Worksheet.CellDirection.Disabled, 2, 7, 2, 7)]
        public void CurrentCellDirectionTest(Worksheet.CellDirection direction, int givenInitialColumn, int givenInitialRow, int expectedColumn, int expectedRow )
        {
            Worksheet worksheet = new Worksheet();
            worksheet.CurrentCellDirection = direction;
            worksheet.SetCurrentCellAddress(givenInitialColumn, givenInitialRow);
            Assert.Equal(givenInitialRow, worksheet.GetCurrentRowNumber());
            Assert.Equal(givenInitialColumn, worksheet.GetCurrentColumnNumber());
            worksheet.AddNextCell("test");
            Assert.Equal(expectedRow, worksheet.GetCurrentRowNumber());
            Assert.Equal(expectedColumn, worksheet.GetCurrentColumnNumber());
        }

        [Theory(DisplayName = "Test of the DefaultColumnWidth property")]
        [InlineData(1f)]
        [InlineData(15.5f)]
        [InlineData(0f)]
        [InlineData(255f)]
        public void DefaultColumnWidthTest(float value)
        {
            Worksheet worksheet = new Worksheet();
            Assert.Equal(Worksheet.DEFAULT_COLUMN_WIDTH, worksheet.DefaultColumnWidth);
            worksheet.DefaultColumnWidth = value;
            Assert.Equal(value, worksheet.DefaultColumnWidth);
        }

        [Theory(DisplayName = "Test of the failing DefaultColumnWidth property")]
        [InlineData(-1f)]
        [InlineData(255.1f)]
        public void DefaultColumnWidthTest2(float value)
        {
            Worksheet worksheet = new Worksheet();
            Assert.Throws<NanoXLSX.Exceptions.RangeException>(() => worksheet.DefaultColumnWidth = value);
        }

        [Theory(DisplayName = "Test of the DefaultRowHeight property")]
        [InlineData(1f)]
        [InlineData(15.5f)]
        [InlineData(0f)]
        [InlineData(409.5)]
        public void DefaultRowHeightTest(float value)
        {
            Worksheet worksheet = new Worksheet();
            Assert.Equal(Worksheet.DEFAULT_ROW_HEIGHT, worksheet.DefaultRowHeight);
            worksheet.DefaultRowHeight = value;
            Assert.Equal(value, worksheet.DefaultRowHeight);
        }

        [Theory(DisplayName = "Test of the failing DefaultRowHeight property")]
        [InlineData(-1f)]
        [InlineData(410f)]
        public void DefaultRowHeightTest2(float value)
        {
            Worksheet worksheet = new Worksheet();
            Assert.Throws<NanoXLSX.Exceptions.RangeException>(() => worksheet.DefaultRowHeight = value);
        }

        [Fact(DisplayName = "Test of the get function of the HiddenRows property")]
        public void HiddenRowsTest()
        {
            Worksheet worksheet = new Worksheet();
            Assert.NotNull(worksheet.HiddenRows);
            Assert.Empty(worksheet.HiddenRows);
            worksheet.AddHiddenRow(2);
            worksheet.AddHiddenRow(5);
            Assert.Equal(2, worksheet.HiddenRows.Count);
            Assert.Contains(worksheet.HiddenRows, item => (item.Key.Equals(2) && item.Value.Equals(true)));
            Assert.Contains(worksheet.HiddenRows, item => (item.Key.Equals(5) && item.Value.Equals(true)));
            worksheet.RemoveHiddenRow(2);
            Assert.Single(worksheet.HiddenRows);
            Assert.Contains(worksheet.HiddenRows, item => (item.Key.Equals(5) && item.Value.Equals(true)));
        }

        [Fact(DisplayName = "Test of the get function of the MergedCells property")]
        public void MergedCellsTest()
        {
            Worksheet worksheet = new Worksheet();
            Assert.NotNull(worksheet.MergedCells);
            Assert.Empty(worksheet.MergedCells);
            NanoXLSX.Range range1 = new NanoXLSX.Range("A2:C3");
            NanoXLSX.Range range2 = new NanoXLSX.Range("S3:R2");
            worksheet.MergeCells(range1);
            worksheet.MergeCells(range2);
            Assert.Equal(2, worksheet.MergedCells.Count);
            Assert.Contains(worksheet.MergedCells, item => (item.Key.Equals("A2:C3") && item.Value.Equals(range1)));
            Assert.Contains(worksheet.MergedCells, item => (item.Key.Equals("R2:S3") && item.Value.Equals(range2)));
            worksheet.RemoveMergedCells(range1.ToString());
            Assert.Single(worksheet.MergedCells);
            Assert.Contains(worksheet.MergedCells, item => (item.Key.Equals("R2:S3") && item.Value.ToString().Equals("R2:S3")));
        }


        [Fact(DisplayName = "Test of the get function of the RowHeights property")]
        public void RowHeightsTest()
        {
            Worksheet worksheet = new Worksheet();
            Assert.NotNull(worksheet.RowHeights);
            Assert.Empty(worksheet.RowHeights);
            worksheet.SetRowHeight(2, 15.3f);
            worksheet.SetRowHeight(5, 100f);
            Assert.Equal(2, worksheet.RowHeights.Count);
            Assert.Contains(worksheet.RowHeights, item => (item.Key.Equals(2) && item.Value.Equals(15.3f)));
            Assert.Contains(worksheet.RowHeights, item => (item.Key.Equals(5) && item.Value.Equals(100f)));
            worksheet.RemoveRowHeight(2);
            Assert.Single(worksheet.RowHeights);
            Assert.Contains(worksheet.RowHeights, item => (item.Key.Equals(5) && item.Value.Equals(100f)));
        }

        [Fact(DisplayName = "Test of the get function of the SelectedCells property")]
        public void SelectedCellsTest()
        {
            Worksheet worksheet = new Worksheet();
            Assert.Null(worksheet.SelectedCells);
            worksheet.SetSelectedCells("B2:D4");
            NanoXLSX.Range ExpectedRange = new NanoXLSX.Range("B2:D4");
            Assert.Equal(ExpectedRange, worksheet.SelectedCells);
            worksheet.RemoveSelectedCells();
            Assert.Null(worksheet.SelectedCells);
        }

        [Fact(DisplayName = "Test of the SheetID property, as well as failing if invalid")]
        public void SheetIDTest()
        {
            Worksheet worksheet = new Worksheet();
            Assert.Equal(0, worksheet.SheetID);
            worksheet.SheetID = 12;
            Assert.Equal(12, worksheet.SheetID);
            Assert.Throws<NanoXLSX.Exceptions.FormatException>(() => worksheet.SheetID = 0);
            Assert.Throws<NanoXLSX.Exceptions.FormatException>(() => worksheet.SheetID = -1);
        }

        [Theory(DisplayName = "Test of the  SheetName property")]
        [InlineData(".")]
        [InlineData(" ")]
        [InlineData("Test")]
        [InlineData("xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx")]
        public void NameTest(string name)
        {
            Worksheet worksheet = new Worksheet();
            Assert.Null(worksheet.SheetName);
            worksheet.SheetName = name;
            Assert.Equal(name, worksheet.SheetName);
        }

        [Theory(DisplayName = "Test failing of the set functions of the SheetName property if a worksheet name is invalid")]
        [InlineData(null)]
        [InlineData("")]
        [InlineData("xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx")]
        [InlineData("A[B")]
        [InlineData("A]B")]
        [InlineData("A*B")]
        [InlineData("A?B")]
        [InlineData("A/B")]
        [InlineData("A\\B")]
        public void NameFailTest(string name)
        {
            Worksheet worksheet = new Worksheet();
            Exception ex = Assert.Throws<NanoXLSX.Exceptions.FormatException>(() => worksheet.SheetName = name);
            Assert.Equal(typeof(NanoXLSX.Exceptions.FormatException), ex.GetType());
        }

        [Theory(DisplayName = "Test of the get function of the SheetProtectionPassword property")]
        [InlineData(null, null)]
        [InlineData("", null)]
        [InlineData(" ", " ")]
        [InlineData("test", "test")]
        public void SheetProtectionPasswordTest(String givenValue, String expectedValue)
        {
            Worksheet worksheet = new Worksheet();
            Assert.Null(worksheet.SheetProtectionPassword);
            worksheet.SetSheetProtectionPassword(givenValue);
            Assert.Equal(expectedValue, worksheet.SheetProtectionPassword);
            worksheet.SetSheetProtectionPassword(null);
            Assert.Null(worksheet.SheetProtectionPassword);
        }

        [Fact(DisplayName = "Test of the SheetProtectionValues property")]
        public void SheetProtectionValuesTest()
        {
            Worksheet worksheet = new Worksheet();
            Assert.NotNull(worksheet.SheetProtectionValues);
            Assert.Empty(worksheet.SheetProtectionValues);
            worksheet.AddAllowedActionOnSheetProtection(Worksheet.SheetProtectionValue.deleteRows);
            worksheet.AddAllowedActionOnSheetProtection(Worksheet.SheetProtectionValue.formatRows);
            Assert.Equal(2, worksheet.SheetProtectionValues.Count);
            Assert.Contains(worksheet.SheetProtectionValues, item => (item.Equals(Worksheet.SheetProtectionValue.deleteRows)));
            Assert.Contains(worksheet.SheetProtectionValues, item => (item.Equals(Worksheet.SheetProtectionValue.formatRows)));
            worksheet.RemoveAllowedActionOnSheetProtection(Worksheet.SheetProtectionValue.deleteRows);
            Assert.Single(worksheet.SheetProtectionValues);
            Assert.Contains(worksheet.SheetProtectionValues, item => (item.Equals(Worksheet.SheetProtectionValue.formatRows)));
        }

        [Fact(DisplayName = "Test of the UseSheetProtection property")]
        public void UseSheetProtectionTest()
        {
            Worksheet worksheet = new Worksheet();
            Assert.False(worksheet.UseSheetProtection);
            worksheet.UseSheetProtection = true;
            Assert.True(worksheet.UseSheetProtection);
        }

        [Fact(DisplayName = "Test of the WorkbookReference property")]
        public void WorkbookReferenceTest()
        {
            Workbook workbook = new Workbook("test.xlsx", "test");
            Worksheet worksheet = new Worksheet();
            Assert.Null(worksheet.WorkbookReference);
            worksheet.WorkbookReference = workbook;
            Assert.NotNull(worksheet.WorkbookReference);
            Assert.Equal("test.xlsx", worksheet.WorkbookReference.Filename);
        }

        [Fact(DisplayName = "Test of the Hidden property")]
        public void HiddenTest()
        {
            Worksheet worksheet = new Worksheet();
            Assert.False(worksheet.Hidden);
            worksheet.Hidden = true;
            Assert.True(worksheet.Hidden);
        }

        [Fact(DisplayName = "Test of the get function of the PaneSplitTopHeight property")]
        public void PaneSplitTopHeightTest()
        {
            Worksheet worksheet = new Worksheet();
            Assert.Null(worksheet.PaneSplitTopHeight);
            worksheet.SetSplit(10f, 22.2f, new Address("A2"), Worksheet.WorksheetPane.bottomLeft);
            Assert.NotNull(worksheet.PaneSplitTopHeight);
            Assert.Equal(22.2f, worksheet.PaneSplitTopHeight);
            worksheet.ResetSplit();
            Assert.Null(worksheet.PaneSplitTopHeight);
        }

        [Fact(DisplayName = "Test of the get function of the PaneSplitLeftWidth property")]
        public void PaneSplitLeftWidthTest()
        {
            Worksheet worksheet = new Worksheet();
            Assert.Null(worksheet.PaneSplitLeftWidth);
            worksheet.SetSplit(11.1f, 20f, new Address("A2"), Worksheet.WorksheetPane.bottomLeft);
            Assert.NotNull(worksheet.PaneSplitLeftWidth);
            Assert.Equal(11.1f, worksheet.PaneSplitLeftWidth);
            worksheet.ResetSplit();
            Assert.Null(worksheet.PaneSplitLeftWidth);
        }

        [Fact(DisplayName = "Test of the get function of the FreezeSplitPanes property")]
        public void FreezeSplitPanesTest()
        {
            Worksheet worksheet = new Worksheet();
            Assert.Null(worksheet.FreezeSplitPanes);
            worksheet.SetSplit(2,2,true, new Address("D4"), Worksheet.WorksheetPane.bottomRight);
            Assert.NotNull(worksheet.FreezeSplitPanes);
            Assert.Equal(true, worksheet.FreezeSplitPanes);
            worksheet.ResetSplit();
            Assert.Null(worksheet.FreezeSplitPanes);
        }

        [Fact(DisplayName = "Test of the get function of the PaneSplitTopLeftCell property")]
        public void PaneSplitTopLeftCellTest()
        {
            Worksheet worksheet = new Worksheet();
            Assert.Null(worksheet.PaneSplitTopLeftCell);
            worksheet.SetSplit(10f, 22.2f, new Address("C4"), Worksheet.WorksheetPane.bottomLeft);
            Assert.NotNull(worksheet.PaneSplitTopLeftCell);
            Assert.Equal("C4", worksheet.PaneSplitTopLeftCell.Value.GetAddress());
            worksheet.ResetSplit();
            Assert.Null(worksheet.PaneSplitTopLeftCell);
        }

        [Fact(DisplayName = "Test of the get function of the PaneSplitAddress property")]
        public void PaneSplitAddressTest()
        {
            Worksheet worksheet = new Worksheet();
            Assert.Null(worksheet.PaneSplitAddress);
            worksheet.SetSplit(2, 2, true, new Address("D4"), Worksheet.WorksheetPane.bottomRight);
            Assert.NotNull(worksheet.PaneSplitAddress);
            Assert.Equal("C3", worksheet.PaneSplitAddress.Value.GetAddress());
            worksheet.ResetSplit();
            Assert.Null(worksheet.PaneSplitAddress);
        }

        [Fact(DisplayName = "Test of the get function of the ActivePane property")]
        public void ActivePaneTest()
        {
            Worksheet worksheet = new Worksheet();
            Assert.Null(worksheet.ActivePane);
            worksheet.SetSplit(2, 2, true, new Address("D4"), Worksheet.WorksheetPane.bottomRight);
            Assert.NotNull(worksheet.ActivePane);
            Assert.Equal(Worksheet.WorksheetPane.bottomRight, worksheet.ActivePane);
            worksheet.ResetSplit();
            Assert.Null(worksheet.ActivePane);
        }

        [Fact(DisplayName = "Test of the get function of the ActiveStyle property")]
        public void ActiveStyleTest()
        {
            Worksheet worksheet = new Worksheet();
            Assert.Null(worksheet.ActiveStyle);
            worksheet.SetActiveStyle(BasicStyles.DottedFill_0_125);
            Assert.NotNull(worksheet.ActiveStyle);
            Assert.True(BasicStyles.DottedFill_0_125.Equals(worksheet.ActiveStyle));
            worksheet.ClearActiveStyle();
            Assert.Null(worksheet.ActiveStyle);
        }

        public static Worksheet InitWorksheet(Worksheet worksheet, string address, Worksheet.CellDirection direction, Style style = null)
        {
            if (worksheet == null)
            {
                worksheet = new Worksheet();
            }
            worksheet.SetCurrentCellAddress(address);
            worksheet.CurrentCellDirection = direction;
            if (style != null)
            {
                worksheet.SetActiveStyle(style);
            }
            return worksheet;
        }

        public static void AssertAddedCell(Worksheet worksheet, int numberOfEntries, string expectedAddress, Cell.CellType expectedType, Style expectedStyle, object expectedValue, int nextColumn, int nextRow)
        {
            Assert.Equal(numberOfEntries, worksheet.Cells.Count);
            Assert.Contains(worksheet.Cells, cell => cell.Key.Equals(expectedAddress));
            Assert.Equal(expectedType, worksheet.Cells[expectedAddress].DataType);
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

        private void AssertConstructorBasics(Worksheet worksheet)
        {
            Assert.NotNull(worksheet);
            Assert.NotNull(worksheet.Cells);
            Assert.Empty(worksheet.Cells);
            Assert.Equal(0, worksheet.GetCurrentRowNumber());
            Assert.Equal(0, worksheet.GetCurrentColumnNumber());
            Assert.Equal(Worksheet.DEFAULT_COLUMN_WIDTH, worksheet.DefaultColumnWidth);
            Assert.Equal(Worksheet.DEFAULT_ROW_HEIGHT, worksheet.DefaultRowHeight);
            Assert.NotNull(worksheet.RowHeights);
            Assert.Empty(worksheet.RowHeights);
            Assert.NotNull(worksheet.MergedCells);
            Assert.Empty(worksheet.MergedCells);
            Assert.NotNull(worksheet.SheetProtectionValues);
            Assert.Empty(worksheet.SheetProtectionValues);
            Assert.NotNull(worksheet.HiddenRows);
            Assert.Empty(worksheet.HiddenRows);
            Assert.NotNull(worksheet.Columns);
            Assert.Empty(worksheet.Columns);
            Assert.Null(worksheet.ActiveStyle);
            Assert.Null(worksheet.ActivePane);
        }

    }
}
