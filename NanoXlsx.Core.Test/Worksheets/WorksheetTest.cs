using System;
using System.Collections.Generic;
using System.IO;
using NanoXLSX.Shared.Exceptions;
using NanoXLSX.Styles;
using Xunit;
using static NanoXLSX.Worksheet;
using FormatException = NanoXLSX.Shared.Exceptions.FormatException;

namespace NanoXLSX.Test.Worksheets
{
    // Ensure that these tests are executed sequentially, since static repository methods may be called 
    [Collection(nameof(SequentialCollection))]
    public class WorksheetTest
    {
        public enum RangeRepresentation
        {
            StringExpression,
            RangeObject,
            Addresses
        }

        [Fact(DisplayName = "Test of the default constructor")]
        public void ConstructorTest()
        {
            Worksheet worksheet = new Worksheet();
            AssertConstructorBasics(worksheet);
            Assert.Null(worksheet.WorkbookReference);
            Assert.Equal(0, worksheet.SheetID);
        }

        [Theory(DisplayName = "Test of the constructor with the worksheet name")]
        [InlineData(".")]
        [InlineData(" ")]
        [InlineData("Test")]
        [InlineData("xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx")]
        public void ConstructorTest2(string name)
        {
            Worksheet worksheet = new Worksheet(name);
            AssertConstructorBasics(worksheet);
            Assert.Null(worksheet.WorkbookReference);
            Assert.Equal(name, worksheet.SheetName);
        }

        [Theory(DisplayName = "Test of the constructor with all parameters")]
        [InlineData(".", 1)]
        [InlineData(" ", 2)]
        [InlineData("Test", 10)]
        [InlineData("xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx", 255)]
        public void ConstructorTest3(string name, int id)
        {
            Workbook workbook = new Workbook("test.xlsx", "sheet2");
            Worksheet worksheet = new Worksheet(name, id, workbook);
            AssertConstructorBasics(worksheet);
            Assert.NotNull(worksheet.WorkbookReference);
            Assert.Equal("test.xlsx", worksheet.WorkbookReference.Filename);
            Assert.Equal(id, worksheet.SheetID);
        }

        [Theory(DisplayName = "Test of the failing constructor if provided with invalid worksheet names")]
        [InlineData("")]
        [InlineData(null)]
        [InlineData("[")]
        [InlineData("................................")]
        public void ConstructorFailingTest(string name)
        {
            Assert.Throws<NanoXLSX.Shared.Exceptions.FormatException>(() => new Worksheet(name));
        }


        [Theory(DisplayName = "Test of the failing constructor if provided with invalid values")]
        [InlineData("", 1)]
        [InlineData(null, 1)]
        [InlineData("[", 1)]
        [InlineData("................................", 0)]
        [InlineData("Test", 0)]
        [InlineData("Test", -1)]
        public void ConstructorFailingTest2(string name, int id)
        {
            Workbook workbook = new Workbook("test.xlsx", "sheet2");
            Assert.Throws<NanoXLSX.Shared.Exceptions.FormatException>(() => new Worksheet(name, id, workbook));
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
        public void CurrentCellDirectionTest(Worksheet.CellDirection direction, int givenInitialColumn, int givenInitialRow, int expectedColumn, int expectedRow)
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
            Assert.Throws<NanoXLSX.Shared.Exceptions.RangeException>(() => worksheet.DefaultColumnWidth = value);
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
            Assert.Throws<NanoXLSX.Shared.Exceptions.RangeException>(() => worksheet.DefaultRowHeight = value);
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
            Assert.Empty(worksheet.SelectedCells);
            worksheet.AddSelectedCells("B2:D4");
            NanoXLSX.Range expectedRange = new NanoXLSX.Range("B2:D4");
            Assert.Contains(expectedRange, worksheet.SelectedCells);
            worksheet.ClearSelectedCells();
            Assert.Empty(worksheet.SelectedCells);
        }

        [Fact(DisplayName = "Test of the SheetID property, as well as failing if invalid")]
        public void SheetIDTest()
        {
            Worksheet worksheet = new Worksheet();
            Assert.Equal(0, worksheet.SheetID);
            worksheet.SheetID = 12;
            Assert.Equal(12, worksheet.SheetID);
            Assert.Throws<NanoXLSX.Shared.Exceptions.FormatException>(() => worksheet.SheetID = 0);
            Assert.Throws<NanoXLSX.Shared.Exceptions.FormatException>(() => worksheet.SheetID = -1);
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
            Exception ex = Assert.Throws<NanoXLSX.Shared.Exceptions.FormatException>(() => worksheet.SheetName = name);
            Assert.Equal(typeof(NanoXLSX.Shared.Exceptions.FormatException), ex.GetType());
        }

        [Theory(DisplayName = "Test of the get function of the SheetProtectionPassword property")]
        [InlineData(null, null)]
        [InlineData("", null)]
        [InlineData(" ", " ")]
        [InlineData("test", "test")]
        public void SheetProtectionPasswordTest(string givenValue, string expectedValue)
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

        [Fact(DisplayName = "Test of the failing set function of the Hidden property when trying to hide all worksheets")]
        public void HiddenFailTest()
        {
            Workbook workbook = new Workbook("test1");
            workbook.AddWorksheet("test2");
            workbook.Worksheets[1].Hidden = true;
            Assert.False(workbook.Worksheets[0].Hidden);
            Assert.True(workbook.Worksheets[1].Hidden);
            Assert.Throws<WorksheetException>(() => workbook.Worksheets[0].Hidden = true);
        }

        [Fact(DisplayName = "Test of the failing set function of the Hidden property when trying to hide all worksheets (scenario with 3 worksheets)")]
        public void HiddenFailTest2()
        {
            Workbook workbook = new Workbook("test1");
            workbook.AddWorksheet("test2");
            workbook.AddWorksheet("test3");
            workbook.SetSelectedWorksheet(1);
            workbook.Worksheets[0].Hidden = true;
            workbook.Worksheets[2].Hidden = true;
            Assert.True(workbook.Worksheets[0].Hidden);
            Assert.False(workbook.Worksheets[1].Hidden);
            Assert.True(workbook.Worksheets[2].Hidden);
            Assert.Throws<WorksheetException>(() => workbook.Worksheets[1].Hidden = true);
        }

        [Fact(DisplayName = "Test of the failing set function of the Hidden property when trying to hide all worksheets by adding hidden worksheets to a workbook")]
        public void HiddenFailTest3()
        {
            Worksheet hidden = new Worksheet("test1");
            hidden.Hidden = true;
            Workbook workbook = new Workbook();
            Assert.Empty(workbook.Worksheets);
            Assert.Throws<WorksheetException>(() => workbook.AddWorksheet(hidden));
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

        [Fact(DisplayName = "Test of the RemoveCell function with column and row")]
        public void RemoveCelltest()
        {
            Worksheet worksheet = new Worksheet();
            List<string> values = new List<string> { "test1", "test2", "test3" };
            worksheet.AddCellRange(values, "A1:A3");
            Assert.Equal(3, worksheet.Cells.Count);
            bool result = worksheet.RemoveCell(0, 1);
            Assert.True(result);
            Assert.Equal(2, worksheet.Cells.Count);
            Assert.DoesNotContain(worksheet.Cells, cell => cell.Key.Equals("A2"));
            result = worksheet.RemoveCell(0, 1); // re-test
            Assert.False(result);
            Assert.Equal(2, worksheet.Cells.Count);
        }

        [Fact(DisplayName = "Test of the RemoveCell function with address")]
        public void RemoveCellTest2()
        {
            Worksheet worksheet = new Worksheet();
            List<string> values = new List<string> { "test1", "test2", "test3" };
            worksheet.AddCellRange(values, "A1:A3");
            Assert.Equal(3, worksheet.Cells.Count);
            bool result = worksheet.RemoveCell("A3");
            Assert.True(result);
            Assert.Equal(2, worksheet.Cells.Count);
            Assert.DoesNotContain(worksheet.Cells, cell => cell.Key.Equals("A3"));
            result = worksheet.RemoveCell("A3"); // re-test
            Assert.False(result);
            Assert.Equal(2, worksheet.Cells.Count);
        }

        [Fact(DisplayName = "Test of the RemoveCell function when no cells are defined")]
        public void RemoveCellTest3()
        {
            Worksheet worksheet = new Worksheet();
            Assert.Empty(worksheet.Cells);
            bool result = worksheet.RemoveCell(2, 5);
            Assert.False(result);
            result = worksheet.RemoveCell("A3");
            Assert.False(result);
        }

        [Theory(DisplayName = "Test of the AddAllowedActionOnSheetProtection function")]
        [InlineData(SheetProtectionValue.deleteRows, 1, null)]
        [InlineData(SheetProtectionValue.formatRows, 1, null)]
        [InlineData(SheetProtectionValue.selectLockedCells, 2, SheetProtectionValue.selectUnlockedCells)]
        [InlineData(SheetProtectionValue.selectUnlockedCells, 1, null)]
        [InlineData(SheetProtectionValue.autoFilter, 1, null)]
        [InlineData(SheetProtectionValue.sort, 1, null)]
        [InlineData(SheetProtectionValue.insertRows, 1, null)]
        [InlineData(SheetProtectionValue.deleteColumns, 1, null)]
        [InlineData(SheetProtectionValue.formatCells, 1, null)]
        [InlineData(SheetProtectionValue.formatColumns, 1, null)]
        [InlineData(SheetProtectionValue.insertHyperlinks, 1, null)]
        [InlineData(SheetProtectionValue.insertColumns, 1, null)]
        [InlineData(SheetProtectionValue.objects, 1, null)]
        [InlineData(SheetProtectionValue.pivotTables, 1, null)]
        [InlineData(SheetProtectionValue.scenarios, 1, null)]

        public void AddAllowedActionOnSheetProtectionTest(SheetProtectionValue typeOfProtection, int expectedSize, SheetProtectionValue? additionalExpectedValue)
        {
            Worksheet worksheet = new Worksheet();
            Assert.False(worksheet.UseSheetProtection);
            Assert.Empty(worksheet.SheetProtectionValues);
            worksheet.AddAllowedActionOnSheetProtection(typeOfProtection);
            Assert.Contains(worksheet.SheetProtectionValues, item => item == typeOfProtection);
            if (additionalExpectedValue != null)
            {
                Assert.Contains(worksheet.SheetProtectionValues, item => item == additionalExpectedValue);
            }
            Assert.Equal(expectedSize, worksheet.SheetProtectionValues.Count);
            worksheet.AddAllowedActionOnSheetProtection(typeOfProtection); // Should not lead to an additional value
            Assert.Equal(expectedSize, worksheet.SheetProtectionValues.Count);
            SheetProtectionValue additionalValue;
            if (typeOfProtection == SheetProtectionValue.objects)
            {
                additionalValue = SheetProtectionValue.sort;
            }
            else
            {
                additionalValue = SheetProtectionValue.objects;
            }
            worksheet.AddAllowedActionOnSheetProtection(additionalValue);
            Assert.Contains(worksheet.SheetProtectionValues, item => item == additionalValue);
            Assert.Equal(expectedSize + 1, worksheet.SheetProtectionValues.Count);
            Assert.True(worksheet.UseSheetProtection);
        }

        [Theory(DisplayName = "Test of the GetCell function with an Address object")]
        [InlineData("C2", "test", "C2")]
        [InlineData("C1,C2,C3", 22, "C2")]
        [InlineData("A1,B1,C1,D1", true, "C1")]
        public void GetCellTest(string definedCells, object definedSample, string expectedAddress)
        {
            List<string> addresses = TestUtils.SplitValuesAsList(definedCells);
            Worksheet worksheet = new Worksheet();
            foreach (string address in addresses)
            {
                worksheet.AddCell(definedSample, address);
            }
            Cell cell = worksheet.GetCell(new Address(expectedAddress));
            Assert.NotNull(cell);
            Assert.Equal(definedSample, cell.Value);
            Assert.Equal(expectedAddress, cell.CellAddress);
        }

        [Theory(DisplayName = "Test of the GetCell function with column and row")]
        [InlineData("C2", "test", 2, 1)]
        [InlineData("C1,C2,C3", 22, 2, 1)]
        [InlineData("A1,B1,C1,D1", true, 2, 0)]
        public void GetCellTest2(string definedCells, object definedSample, int expectedColumn, int expectedRow)
        {
            List<string> addresses = TestUtils.SplitValuesAsList(definedCells);
            Worksheet worksheet = new Worksheet();
            foreach (string address in addresses)
            {
                worksheet.AddCell(definedSample, address);
            }
            Cell cell = worksheet.GetCell(expectedColumn, expectedRow);
            Assert.NotNull(cell);
            Assert.Equal(definedSample, cell.Value);
            Assert.Equal(new Address(expectedColumn, expectedRow), cell.CellAddress2);
        }

        [Theory(DisplayName = "Test of the failing GetCell function with an Address object")]
        [InlineData("", null, "C2")]
        [InlineData("C1,C2,C3", 22, "D2")]
        public void GetCellFailTest(string definedCells, object definedSample, string expectedAddress)
        {
            List<string> addresses = TestUtils.SplitValuesAsList(definedCells);
            Worksheet worksheet = new Worksheet();
            foreach (string address in addresses)
            {
                worksheet.AddCell(definedSample, address);
            }
            Assert.Throws<WorksheetException>(() => worksheet.GetCell(new Address(expectedAddress)));
        }

        [Theory(DisplayName = "Test of the failing GetCell function with a column and row")]
        [InlineData("", null, 2, 1, typeof(WorksheetException))]
        [InlineData("C1,C2,C3", 22, 3, 1, typeof(WorksheetException))]
        [InlineData("C1,C2,C3", 22, -1, 2, typeof(RangeException))]
        [InlineData("C1,C2,C3", 22, 2, -1, typeof(RangeException))]
        [InlineData("C1,C2,C3", 22, 16384, 2, typeof(RangeException))]
        [InlineData("C1,C2,C3", 22, 2, 1048576, typeof(RangeException))]
        public void GetCellFailTest2(string definedCells, object definedSample, int expectedColumn, int expectedRow, Type exceptionType)
        {
            List<string> addresses = TestUtils.SplitValuesAsList(definedCells);
            Worksheet worksheet = new Worksheet();
            foreach (string address in addresses)
            {
                worksheet.AddCell(definedSample, address);
            }
            Exception exception = Assert.ThrowsAny<Exception>(() => worksheet.GetCell(expectedColumn, expectedRow));
            Assert.Equal(exceptionType, exception.GetType());
        }

        [Theory(DisplayName = "Test of the HasCell function with an Address object")]
        [InlineData("C2", "C2", true)]
        [InlineData("C2", "C3", false)]
        [InlineData("", "C2", false)]
        [InlineData("C2,C3,C4", "C2", true)]
        [InlineData("C2,C3,C4", "D2", false)]
        public void HasCellTest(string definedCells, string givenAddress, bool expectedResult)
        {
            List<string> addresses = TestUtils.SplitValuesAsList(definedCells);
            Worksheet worksheet = new Worksheet();
            foreach (string address in addresses)
            {
                worksheet.AddCell("test", address);
            }
            Assert.Equal(expectedResult, worksheet.HasCell(new Address(givenAddress)));
        }

        [Theory(DisplayName = "Test of the HasCell function with a column and row")]
        [InlineData("C2", 2, 1, true)]
        [InlineData("C2", 2, 2, false)]
        [InlineData("", 2, 1, false)]
        [InlineData("C2,C3,C4", 2, 1, true)]
        [InlineData("C2,C3,C4", 3, 1, false)]
        public void HasCellTest2(string definedCells, int givenColumn, int givenRow, bool expectedResult)
        {
            List<string> addresses = TestUtils.SplitValuesAsList(definedCells);
            Worksheet worksheet = new Worksheet();
            foreach (string address in addresses)
            {
                worksheet.AddCell("test", address);
            }
            Assert.Equal(expectedResult, worksheet.HasCell(givenColumn, givenRow));
        }

        [Theory(DisplayName = "Test of the failing HasCell function with a column and row")]
        [InlineData(-1, 2)]
        [InlineData(2, -1)]
        [InlineData(16384, 2)]
        [InlineData(2, 1048576)]
        public void HasCellFailTest(int givenColumn, int givenRow)
        {
            Worksheet worksheet = new Worksheet();
            worksheet.AddCell("test", "C3");
            Assert.Throws<RangeException>(() => worksheet.GetCell(givenColumn, givenRow));
        }

        [Theory(DisplayName = "Test of the GetLastCellAddress function with an empty worksheet")]
        [InlineData(false, false, false)]
        [InlineData(false, false, true)]
        [InlineData(false, true, true)]
        [InlineData(false, true, false)]
        [InlineData(true, false, false)]
        public void GetLastCellAddressTest(bool hasColumns, bool hasHiddenRows, bool hasRowHeights)
        {
            Worksheet worksheet = new Worksheet();
            if (hasColumns)
            {
                worksheet.AddHiddenColumn(0);
                worksheet.AddHiddenColumn(1);
                worksheet.AddHiddenColumn(2);
            }
            if (hasHiddenRows)
            {
                worksheet.AddHiddenRow(0);
                worksheet.AddHiddenRow(1);
                worksheet.AddHiddenRow(2);
            }
            if (hasRowHeights)
            {
                worksheet.SetRowHeight(0, 22.2f);
                worksheet.SetRowHeight(1, 22.2f);
                worksheet.SetRowHeight(2, 22.2f);
            }
            Address? address = worksheet.GetLastCellAddress();
            Assert.Null(address);
        }

        [Fact(DisplayName = "Test of the GetLastCellAddress function with an empty worksheet but defined columns and rows")]
        public void GetLastCellAddressTest2()
        {
            Worksheet worksheet = new Worksheet();
            worksheet.AddHiddenColumn(0);
            worksheet.AddHiddenColumn(1);
            worksheet.AddHiddenColumn(2);
            worksheet.AddHiddenRow(0);
            worksheet.AddHiddenRow(1);
            Address? address = worksheet.GetLastCellAddress();
            Assert.NotNull(address);
            Assert.Equal("C2", address.Value.GetAddress());
        }

        [Fact(DisplayName = "Test of the GetLastCellAddress function with an empty worksheet but defined columns and rows with gaps")]
        public void GetLastCellAddressTest3()
        {
            Worksheet worksheet = new Worksheet();
            worksheet.AddHiddenColumn(0);
            worksheet.AddHiddenColumn(1);
            worksheet.AddHiddenColumn(10);
            worksheet.AddHiddenRow(0);
            worksheet.AddHiddenRow(1);
            worksheet.SetRowHeight(10, 22.2f);
            Address? address = worksheet.GetLastCellAddress();
            Assert.NotNull(address);
            Assert.Equal("K11", address.Value.GetAddress());
        }

        [Fact(DisplayName = "Test of the GetLastCellAddress function with defined columns and rows where cells are defined below the last column and row")]
        public void GetLastCellAddressTest4()
        {
            Worksheet worksheet = new Worksheet();
            worksheet.AddHiddenColumn(0);
            worksheet.AddHiddenColumn(1);
            worksheet.AddHiddenColumn(10);
            worksheet.AddHiddenRow(0);
            worksheet.AddHiddenRow(1);
            worksheet.SetRowHeight(10, 22.2f);
            worksheet.AddCell("test", "E5");
            Address? address = worksheet.GetLastCellAddress();
            Assert.NotNull(address);
            Assert.Equal("K11", address.Value.GetAddress());
        }

        [Fact(DisplayName = "Test of the GetLastCellAddress function with defined columns and rows where cells are defined above the last column and row")]
        public void GetLastCellAddressTest5()
        {
            Worksheet worksheet = new Worksheet();
            worksheet.AddHiddenColumn(0);
            worksheet.AddHiddenColumn(1);
            worksheet.AddHiddenColumn(10);
            worksheet.AddHiddenRow(0);
            worksheet.AddHiddenRow(1);
            worksheet.SetRowHeight(10, 22.2f);
            worksheet.AddCell("test", "L12");
            Address? address = worksheet.GetLastCellAddress();
            Assert.NotNull(address);
            Assert.Equal("L12", address.Value.GetAddress());
        }

        [Theory(DisplayName = "Test of the GetLastDataCellAddress function with an empty worksheet")]
        [InlineData(false, false, false)]
        [InlineData(false, false, true)]
        [InlineData(false, true, true)]
        [InlineData(false, true, false)]
        [InlineData(true, false, false)]
        [InlineData(true, true, false)]
        [InlineData(true, false, true)]
        [InlineData(true, true, true)]
        public void GetLastDataCellAddressTest(bool hasColumns, bool hasHiddenRows, bool hasRowHeights)
        {
            Worksheet worksheet = new Worksheet();
            if (hasColumns)
            {
                worksheet.AddHiddenColumn(0);
                worksheet.AddHiddenColumn(1);
                worksheet.AddHiddenColumn(2);
            }
            if (hasHiddenRows)
            {
                worksheet.AddHiddenRow(0);
                worksheet.AddHiddenRow(1);
                worksheet.AddHiddenRow(2);
            }
            if (hasRowHeights)
            {
                worksheet.SetRowHeight(0, 22.2f);
                worksheet.SetRowHeight(1, 22.2f);
                worksheet.SetRowHeight(2, 22.2f);
            }
            Address? address = worksheet.GetLastDataCellAddress();
            Assert.Null(address);
        }

        [Fact(DisplayName = "Test of the GetLastDataCellAddress function with defined columns and rows where cells are defined below the last column and row")]
        public void GetLastDataCellAddressTest2()
        {
            Worksheet worksheet = new Worksheet();
            worksheet.AddHiddenColumn(0);
            worksheet.AddHiddenColumn(1);
            worksheet.AddHiddenColumn(10);
            worksheet.AddHiddenRow(0);
            worksheet.AddHiddenRow(1);
            worksheet.SetRowHeight(10, 22.2f);
            worksheet.AddCell("test", "E7");
            Address? address = worksheet.GetLastDataCellAddress();
            Assert.NotNull(address);
            Assert.Equal("E7", address.Value.GetAddress());
        }

        [Fact(DisplayName = "Test of the GetLastDataCellAddress function with defined columns and rows where cells are defined above the last column and row")]
        public void GetLastDataCellAddressTest3()
        {
            Worksheet worksheet = new Worksheet();
            worksheet.AddHiddenColumn(0);
            worksheet.AddHiddenColumn(1);
            worksheet.AddHiddenColumn(10);
            worksheet.AddHiddenRow(0);
            worksheet.AddHiddenRow(1);
            worksheet.SetRowHeight(10, 22.2f);
            worksheet.AddCell("test", "L12");
            Address? address = worksheet.GetLastDataCellAddress();
            Assert.NotNull(address);
            Assert.Equal("L12", address.Value.GetAddress());
        }

        [Theory(DisplayName = "Test of the GetFirstCellAddress function with an empty worksheet")]
        [InlineData(false, false, false)]
        [InlineData(false, false, true)]
        [InlineData(false, true, true)]
        [InlineData(false, true, false)]
        [InlineData(true, false, false)]
        public void GetFirstCellAddressTest(bool hasColumns, bool hasHiddenRows, bool hasRowHeights)
        {
            Worksheet worksheet = new Worksheet();
            if (hasColumns)
            {
                worksheet.AddHiddenColumn(1);
                worksheet.AddHiddenColumn(2);
                worksheet.AddHiddenColumn(3);
            }
            if (hasHiddenRows)
            {
                worksheet.AddHiddenRow(1);
                worksheet.AddHiddenRow(2);
                worksheet.AddHiddenRow(3);
            }
            if (hasRowHeights)
            {
                worksheet.SetRowHeight(1, 22.2f);
                worksheet.SetRowHeight(2, 22.2f);
                worksheet.SetRowHeight(3, 22.2f);
            }
            Address? address = worksheet.GetFirstCellAddress();
            Assert.Null(address);
        }

        [Fact(DisplayName = "Test of the GetFirstCellAddress function with an empty worksheet but defined columns and rows")]
        public void GetFirstCellAddressTest2()
        {
            Worksheet worksheet = new Worksheet();
            worksheet.AddHiddenColumn(2);
            worksheet.AddHiddenColumn(3);
            worksheet.AddHiddenColumn(4);
            worksheet.AddHiddenRow(1);
            worksheet.AddHiddenRow(2);
            Address? address = worksheet.GetFirstCellAddress();
            Assert.NotNull(address);
            Assert.Equal("C2", address.Value.GetAddress());
        }

        [Fact(DisplayName = "Test of the GetFirstCellAddress function with an empty worksheet but defined columns and rows with gaps")]
        public void GetFirstCellAddressTest3()
        {
            Worksheet worksheet = new Worksheet();
            worksheet.AddHiddenColumn(2);
            worksheet.AddHiddenColumn(3);
            worksheet.AddHiddenColumn(10);
            worksheet.AddHiddenRow(3);
            worksheet.AddHiddenRow(4);
            worksheet.SetRowHeight(10, 22.2f);
            Address? address = worksheet.GetFirstCellAddress();
            Assert.NotNull(address);
            Assert.Equal("C4", address.Value.GetAddress());
        }

        [Fact(DisplayName = "Test of the GetFirstCellAddress function with defined columns and rows where cells are defined above the first column and row")]
        public void GetFirstCellAddressTest4()
        {
            Worksheet worksheet = new Worksheet();
            worksheet.AddHiddenColumn(3);
            worksheet.AddHiddenColumn(4);
            worksheet.AddHiddenColumn(10);
            worksheet.AddHiddenRow(4);
            worksheet.AddHiddenRow(5);
            worksheet.SetRowHeight(10, 22.2f);
            worksheet.AddCell("test", "R5");
            worksheet.AddCell("test", "F11");
            Address? address = worksheet.GetFirstCellAddress();
            Assert.NotNull(address);
            Assert.Equal("D5", address.Value.GetAddress());
        }

        [Fact(DisplayName = "Test of the GetFirstCellAddress function with defined columns and rows where cells are defined below the first column and row")]
        public void GetFirstCellAddressTest5()
        {
            Worksheet worksheet = new Worksheet();
            worksheet.AddHiddenColumn(2);
            worksheet.AddHiddenColumn(4);
            worksheet.AddHiddenColumn(10);
            worksheet.AddHiddenRow(3);
            worksheet.AddHiddenRow(4);
            worksheet.SetRowHeight(10, 22.2f);
            worksheet.AddCell("test", "E5");
            Address? address = worksheet.GetFirstCellAddress();
            Assert.NotNull(address);
            Assert.Equal("C4", address.Value.GetAddress());
        }

        [Theory(DisplayName = "Test of the GetFirstDataCellAddress function with an empty worksheet")]
        [InlineData(false, false, false)]
        [InlineData(false, false, true)]
        [InlineData(false, true, true)]
        [InlineData(false, true, false)]
        [InlineData(true, false, false)]
        [InlineData(true, true, false)]
        [InlineData(true, false, true)]
        [InlineData(true, true, true)]
        public void GetFirstDataCellAddressTest(bool hasColumns, bool hasHiddenRows, bool hasRowHeights)
        {
            Worksheet worksheet = new Worksheet();
            if (hasColumns)
            {
                worksheet.AddHiddenColumn(0);
                worksheet.AddHiddenColumn(1);
                worksheet.AddHiddenColumn(2);
            }
            if (hasHiddenRows)
            {
                worksheet.AddHiddenRow(0);
                worksheet.AddHiddenRow(1);
                worksheet.AddHiddenRow(2);
            }
            if (hasRowHeights)
            {
                worksheet.SetRowHeight(0, 22.2f);
                worksheet.SetRowHeight(1, 22.2f);
                worksheet.SetRowHeight(2, 22.2f);
            }
            Address? address = worksheet.GetFirstDataCellAddress();
            Assert.Null(address);
        }

        [Fact(DisplayName = "Test of the GetFirstDataCellAddress function with defined columns and rows where cells are defined above the first column and row")]
        public void GetFirstDataCellAddressTest2()
        {
            Worksheet worksheet = new Worksheet();
            worksheet.AddHiddenColumn(2);
            worksheet.AddHiddenColumn(3);
            worksheet.AddHiddenColumn(4);
            worksheet.AddHiddenRow(2);
            worksheet.AddHiddenRow(3);
            worksheet.SetRowHeight(4, 22.2f);
            worksheet.AddCell("test", "E6");
            worksheet.AddCell("test", "H9");
            Address? address = worksheet.GetFirstDataCellAddress();
            Assert.NotNull(address);
            Assert.Equal("E6", address.Value.GetAddress());
        }

        [Fact(DisplayName = "Test of the GetFirstDataCellAddress function with defined columns and rows where cells are defined below the first column and row")]
        public void GetFirstDataCellAddressTest3()
        {
            Worksheet worksheet = new Worksheet();
            worksheet.AddHiddenColumn(1);
            worksheet.AddHiddenColumn(2);
            worksheet.AddHiddenColumn(10);
            worksheet.AddHiddenRow(1);
            worksheet.AddHiddenRow(2);
            worksheet.SetRowHeight(10, 22.2f);
            worksheet.AddCell("test", "C5");
            worksheet.AddCell("test", "D7");
            Address? address = worksheet.GetFirstDataCellAddress();
            Assert.NotNull(address);
            Assert.Equal("C5", address.Value.GetAddress());
        }

        [Theory(DisplayName = "Test of the MergeCells function")]
        [InlineData(RangeRepresentation.Addresses, 0, 0, 0, 0, "A1:A1", 1)]
        [InlineData(RangeRepresentation.RangeObject, 1, 1, 1, 1, "B2:B2", 1)]
        [InlineData(RangeRepresentation.StringExpression, 2, 2, 2, 2, "C3:C3", 1)]
        [InlineData(RangeRepresentation.Addresses, 0, 0, 2, 2, "A1:C3", 9)]
        [InlineData(RangeRepresentation.RangeObject, 1, 1, 3, 1, "B2:D2", 3)]
        [InlineData(RangeRepresentation.StringExpression, 2, 2, 2, 4, "C3:C5", 3)]
        [InlineData(RangeRepresentation.Addresses, 2, 2, 0, 0, "C3:A1", 9)]
        [InlineData(RangeRepresentation.StringExpression, 2, 4, 2, 2, "C3:C5", 3)] // String expression is reordered by the test method
        public void MergeCellsTest(RangeRepresentation representation, int givenStartColumn, int givenStartRow, int givenEndColumn, int givenEndRow, string expectedMergedCells, int expectedCount)
        {
            Worksheet worksheet = new Worksheet();
            Address startAddress = new Address(givenStartColumn, givenStartRow);
            Address endAddress = new Address(givenEndColumn, givenEndRow);
            Range range = new Range(startAddress, endAddress);
            Assert.Empty(worksheet.MergedCells);
            string returnedAddress;
            if (representation == RangeRepresentation.Addresses)
            {
                returnedAddress = worksheet.MergeCells(startAddress, endAddress);
            }
            else if (representation == RangeRepresentation.StringExpression)
            {
                returnedAddress = worksheet.MergeCells(range.ToString());
            }
            else
            {
                returnedAddress = worksheet.MergeCells(range);
            }

            Assert.Single(worksheet.MergedCells);
            Assert.Equal(expectedMergedCells, returnedAddress);
            Assert.Contains(worksheet.MergedCells, item => item.Key == expectedMergedCells);
            Assert.Equal(expectedCount, worksheet.MergedCells[expectedMergedCells].ResolveEnclosedAddresses().Count);
        }

        [Theory(DisplayName = "Test of the MergeCells function with more than one range")]
        [InlineData(RangeRepresentation.Addresses, 0, 0, 0, 0, "A1:A1", 1)]
        [InlineData(RangeRepresentation.RangeObject, 1, 1, 1, 1, "B2:B2", 1)]
        [InlineData(RangeRepresentation.StringExpression, 2, 2, 2, 2, "C3:C3", 1)]
        [InlineData(RangeRepresentation.Addresses, 0, 0, 2, 2, "A1:C3", 9)]
        [InlineData(RangeRepresentation.RangeObject, 1, 1, 3, 1, "B2:D2", 3)]
        [InlineData(RangeRepresentation.StringExpression, 2, 2, 2, 4, "C3:C5", 3)]
        [InlineData(RangeRepresentation.Addresses, 2, 2, 0, 0, "C3:A1", 9)]
        [InlineData(RangeRepresentation.StringExpression, 2, 4, 2, 2, "C3:C5", 3)] // String expression is reordered by the test method
        public void MergeCellsTest2(RangeRepresentation representation, int givenStartColumn, int givenStartRow, int givenEndColumn, int givenEndRow, string expectedMergedCells, int expectedCount)
        {
            Worksheet worksheet = new Worksheet();
            Address startAddress = new Address(givenStartColumn, givenStartRow);
            Address endAddress = new Address(givenEndColumn, givenEndRow);
            Range range = new Range(startAddress, endAddress);
            Assert.Empty(worksheet.MergedCells);
            string returnedAddress;
            if (representation == RangeRepresentation.Addresses)
            {
                returnedAddress = worksheet.MergeCells(startAddress, endAddress);
            }
            else if (representation == RangeRepresentation.StringExpression)
            {
                returnedAddress = worksheet.MergeCells(range.ToString());
            }
            else
            {
                returnedAddress = worksheet.MergeCells(range);
            }
            string returnedAddress2 = worksheet.MergeCells("X1:X2");
            Assert.Equal(2, worksheet.MergedCells.Count);
            Assert.Equal(expectedMergedCells, returnedAddress);
            Assert.Contains(worksheet.MergedCells, item => item.Key == expectedMergedCells);
            Assert.Equal(expectedCount, worksheet.MergedCells[expectedMergedCells].ResolveEnclosedAddresses().Count);
            Assert.Contains(worksheet.MergedCells, item => item.Key == returnedAddress2);
            Assert.Equal(2, worksheet.MergedCells[returnedAddress2].ResolveEnclosedAddresses().Count);
        }

        [Fact(DisplayName = "Test of the failing MergeCells function if cell addresses are colliding")]
        public void MergeCellsFailTest()
        {
            Worksheet worksheet = new Worksheet();
            worksheet.MergeCells("A1:D4");
            Assert.Throws<RangeException>(() => worksheet.MergeCells("B4:E4"));
        }

        [Fact(DisplayName = "Test of the failing MergeCells function if the merge range already exists (full intersection)")]
        public void MergeCellsFailTest2()
        {
            Worksheet worksheet = new Worksheet();
            worksheet.MergeCells("A1:D4");
            Assert.Throws<RangeException>(() => worksheet.MergeCells("D4:A1")); // Flip addresses
        }

        [Fact(DisplayName = "Test of the failing MergeCells function if the merge range is invalid (string)")]
        public void MergeCellsFailTest3()
        {
            Worksheet worksheet = new Worksheet();
            Assert.Throws<FormatException>(() => worksheet.MergeCells(""));
        }

        [Fact(DisplayName = "Test of the internal RecalculateAutoFilter function")]
        public void RecalculateAutoFilterTest()
        {
            Workbook workbook = new Workbook(true);
            Worksheet worksheet = workbook.CurrentWorksheet;
            workbook.SaveAsStream(new MemoryStream()); // Dummy call to invoke recalculation
            Assert.Null(worksheet.AutoFilterRange);
            worksheet.AddCell("test", "A100");
            worksheet.AddCell("test", "D50"); // Will expand the range to row 50
            worksheet.AddCell("test", "F2");
            worksheet.SetAutoFilter("B1:E1");
            worksheet.Columns[2].HasAutoFilter = false;
            worksheet.ResetColumn(2);
            workbook.SaveAsStream(new MemoryStream()); // Dummy call to invoke recalculation
            Assert.True(worksheet.Columns[2].HasAutoFilter);
            Assert.Equal("B1:E50", worksheet.AutoFilterRange.ToString());
        }


        [Fact(DisplayName = "Test of the internal RecalculateColumns function")]
        public void RecalculateColumnsTest()
        {
            Workbook workbook = new Workbook(true);
            Worksheet worksheet = workbook.CurrentWorksheet;
            worksheet.SetColumnWidth(1, 22.5f);
            worksheet.SetColumnWidth(2, 22.8f);
            worksheet.AddHiddenColumn(3);
            worksheet.SetColumnWidth(1, Worksheet.DEFAULT_COLUMN_WIDTH); // should not remove the column
            Assert.Equal(3, worksheet.Columns.Count);
            workbook.SaveAsStream(new MemoryStream()); // Dummy call to invoke recalculation
            Assert.Equal(2, worksheet.Columns.Count);
            Assert.False(worksheet.Columns.ContainsKey(1));
        }

        [Fact(DisplayName = "Test of the internal ResolveMergedCells function")]
        public void ResolveMergedCellsTest()
        {
            Workbook workbook = new Workbook(true);
            Worksheet worksheet = workbook.CurrentWorksheet;
            worksheet.AddCell("test", "B1");
            worksheet.AddCell(22.2f, "C1");
            Assert.Equal(2, worksheet.Cells.Count);
            worksheet.MergeCells("B1:D1");
            workbook.SaveAsStream(new MemoryStream()); // Dummy call to invoke resolution
            Assert.Equal(3, worksheet.Cells.Count);
            Assert.Null(worksheet.Cells["B1"].CellStyle);
            Assert.Equal(Cell.CellType.EMPTY, worksheet.Cells["C1"].DataType);
            Assert.True(BasicStyles.MergeCellStyle.Equals(worksheet.Cells["C1"].CellStyle));
            Assert.Equal(22.2f, worksheet.Cells["C1"].Value);
            Assert.True(BasicStyles.MergeCellStyle.Equals(worksheet.Cells["D1"].CellStyle));
            Assert.Equal(Cell.CellType.EMPTY, worksheet.Cells["D1"].DataType);
        }

        [Fact(DisplayName = "Test of the RemoveAutoFilter function")]
        public void RemoveAutoFilterTest()
        {
            Worksheet worksheet = new Worksheet();
            worksheet.SetAutoFilter(1, 5);
            Assert.NotNull(worksheet.AutoFilterRange);
            Assert.Equal("B1:F1", worksheet.AutoFilterRange.Value.ToString());
            worksheet.RemoveAutoFilter();
            Assert.Null(worksheet.AutoFilterRange);
        }

        [Fact(DisplayName = "Test of the RemoveHiddenColumn function")]
        public void RemoveHiddenColumnTest()
        {
            Worksheet worksheet = new Worksheet();
            worksheet.AddHiddenColumn(1);
            worksheet.AddHiddenColumn(2);
            worksheet.AddHiddenColumn(3);
            worksheet.SetColumnWidth(2, 22.2f);
            Assert.Equal(3, worksheet.Columns.Count);
            worksheet.RemoveHiddenColumn(2);
            worksheet.RemoveHiddenColumn(3);
            Assert.Equal(2, worksheet.Columns.Count);
            Assert.False(worksheet.Columns[2].IsHidden);
        }

        [Fact(DisplayName = "Test of the RemoveHiddenColumn function with a string as column expression")]
        public void RemoveHiddenColumnTest2()
        {
            Worksheet worksheet = new Worksheet();
            worksheet.AddHiddenColumn("B");
            worksheet.AddHiddenColumn("C");
            worksheet.AddHiddenColumn("D");
            worksheet.SetColumnWidth(2, 22.2f);
            Assert.Equal(3, worksheet.Columns.Count);
            worksheet.RemoveHiddenColumn("C");
            worksheet.RemoveHiddenColumn("D");
            Assert.Equal(2, worksheet.Columns.Count);
            Assert.False(worksheet.Columns[2].IsHidden);
        }

        [Fact(DisplayName = "Test of the RemoveHiddenRow function")]
        public void RemoveHiddenRowTest()
        {
            Worksheet worksheet = new Worksheet();
            worksheet.AddHiddenRow(1);
            worksheet.AddHiddenRow(2);
            worksheet.AddHiddenRow(3);
            Assert.Equal(3, worksheet.HiddenRows.Count);
            worksheet.RemoveHiddenRow(2);
            Assert.Equal(2, worksheet.HiddenRows.Count);
            Assert.False(worksheet.HiddenRows.ContainsKey(2));
        }

        [Fact(DisplayName = "Test of the RemoveMergedCells function")]
        public void RemoveMergedCellsTest()
        {
            Worksheet worksheet = new Worksheet();
            worksheet.AddCell("test", "B2");
            worksheet.AddCell(22, "B3");
            worksheet.MergeCells("B1:B4");
            Assert.True(worksheet.MergedCells.ContainsKey("B1:B4"));
            worksheet.RemoveMergedCells("B1:B4");
            Assert.Empty(worksheet.MergedCells);
        }

        [Fact(DisplayName = "Test of the RemoveMergedCells function after resolution of merged cells (on save)")]
        public void RemoveMergedCellsTest2()
        {
            Workbook workbook = new Workbook(true);
            Worksheet worksheet = workbook.CurrentWorksheet;
            worksheet.AddCell("test", "B2");
            worksheet.AddCell(22, "B3");
            worksheet.MergeCells("B1:B4");
            Assert.True(worksheet.MergedCells.ContainsKey("B1:B4"));
            workbook.SaveAsStream(new MemoryStream()); // Dummy call to invoke recalculation
            worksheet.RemoveMergedCells("B1:B4");
            Assert.Empty(worksheet.MergedCells);
            Assert.False(BasicStyles.MergeCellStyle.Equals(worksheet.Cells["B2"].CellStyle));
            Assert.False(BasicStyles.MergeCellStyle.Equals(worksheet.Cells["B3"].CellStyle));
            Assert.Equal("test", worksheet.Cells["B2"].Value);
            Assert.Equal(22, worksheet.Cells["B3"].Value);
        }

        [Theory(DisplayName = "Test of the failing RemoveMergedCells function on an invalid range")]
        [InlineData("")]
        [InlineData(null)]
        [InlineData("B1")]
        [InlineData("B1:B5")]
        public void RemoveMergedCellsFailTest(string range)
        {
            Worksheet worksheet = new Worksheet();
            worksheet.AddCell("test", "B2");
            worksheet.AddCell(22, "B3");
            worksheet.MergeCells("B1:B4");
            Assert.True(worksheet.MergedCells.ContainsKey("B1:B4"));
            Assert.Throws<RangeException>(() => worksheet.RemoveMergedCells(range));
        }

        [Fact(DisplayName = "Test of the RemoveSelectedCells function")]
        public void RemoveSelectedCellsTest()
        {
            Worksheet worksheet = new Worksheet();
            Assert.Empty(worksheet.SelectedCells);
            worksheet.AddSelectedCells("B2:D3");
            Assert.Contains(new Range("B2:D3"), worksheet.SelectedCells);
            worksheet.ClearSelectedCells();
            Assert.Empty(worksheet.SelectedCells);
        }

        [Theory(DisplayName = "Test of the RemoveAllowedActionOnSheetProtection function")]
        [InlineData(SheetProtectionValue.deleteRows, SheetProtectionValue.objects, SheetProtectionValue.sort)]
        [InlineData(SheetProtectionValue.formatRows, SheetProtectionValue.objects, SheetProtectionValue.sort)]
        [InlineData(SheetProtectionValue.selectLockedCells, SheetProtectionValue.objects, SheetProtectionValue.sort)]
        [InlineData(SheetProtectionValue.selectUnlockedCells, SheetProtectionValue.objects, SheetProtectionValue.sort)]
        [InlineData(SheetProtectionValue.autoFilter, SheetProtectionValue.objects, SheetProtectionValue.sort)]
        [InlineData(SheetProtectionValue.sort, SheetProtectionValue.objects, SheetProtectionValue.formatRows)]
        [InlineData(SheetProtectionValue.insertRows, SheetProtectionValue.objects, SheetProtectionValue.sort)]
        [InlineData(SheetProtectionValue.deleteColumns, SheetProtectionValue.objects, SheetProtectionValue.sort)]
        [InlineData(SheetProtectionValue.formatCells, SheetProtectionValue.objects, SheetProtectionValue.sort)]
        [InlineData(SheetProtectionValue.formatColumns, SheetProtectionValue.objects, SheetProtectionValue.sort)]
        [InlineData(SheetProtectionValue.insertHyperlinks, SheetProtectionValue.objects, SheetProtectionValue.sort)]
        [InlineData(SheetProtectionValue.insertColumns, SheetProtectionValue.objects, SheetProtectionValue.sort)]
        [InlineData(SheetProtectionValue.objects, SheetProtectionValue.formatColumns, SheetProtectionValue.sort)]
        [InlineData(SheetProtectionValue.pivotTables, SheetProtectionValue.objects, SheetProtectionValue.sort)]
        [InlineData(SheetProtectionValue.scenarios, SheetProtectionValue.objects, SheetProtectionValue.sort)]
        public void RemoveAllowedActionOnSheetProtectionTest(SheetProtectionValue typeOfProtection, SheetProtectionValue additionalValue, SheetProtectionValue notPresentValue)
        {
            Worksheet worksheet = new Worksheet();
            worksheet.AddAllowedActionOnSheetProtection(typeOfProtection);
            worksheet.AddAllowedActionOnSheetProtection(additionalValue);
            int count = worksheet.SheetProtectionValues.Count;
            Assert.True(count >= 2);
            worksheet.RemoveAllowedActionOnSheetProtection(typeOfProtection);
            Assert.Equal(count - 1, worksheet.SheetProtectionValues.Count);
            Assert.DoesNotContain(worksheet.SheetProtectionValues, item => item == typeOfProtection);
            worksheet.RemoveAllowedActionOnSheetProtection(notPresentValue); // should not cause anything
            Assert.Equal(count - 1, worksheet.SheetProtectionValues.Count);
        }

        [Fact(DisplayName = "Test of the SetActiveStyle function")]
        public void SetActiveStyleTest()
        {
            Worksheet worksheet = new Worksheet();
            Assert.Null(worksheet.ActiveStyle);
            worksheet.SetActiveStyle(BasicStyles.Bold);
            Assert.True(BasicStyles.Bold.Equals(worksheet.ActiveStyle));
        }

        [Fact(DisplayName = "Test of the SetActiveStyle function on null")]
        public void SetActiveStyleTest2()
        {
            Worksheet worksheet = new Worksheet();
            Assert.Null(worksheet.ActiveStyle);
            worksheet.SetActiveStyle(null);
            Assert.Null(worksheet.ActiveStyle);
        }

        [Theory(DisplayName = "Test of the SetCurrentCellAddress function with column and row numbers")]
        [InlineData(0, 0)]
        [InlineData(5, 0)]
        [InlineData(0, 5)]
        [InlineData(16383, 1048575)]
        public void SetCurrentCellAddressTest(int column, int row)
        {
            Worksheet worksheet = new Worksheet();
            Assert.Equal(0, worksheet.GetCurrentColumnNumber());
            Assert.Equal(0, worksheet.GetCurrentRowNumber());
            worksheet.GoToNextRow();
            worksheet.GoToNextColumn();
            worksheet.SetCurrentCellAddress(column, row);
            Assert.Equal(column, worksheet.GetCurrentColumnNumber());
            Assert.Equal(row, worksheet.GetCurrentRowNumber());
        }


        [Theory(DisplayName = "Test of the SetCurrentCellAddress function")]
        [InlineData("A1")]
        [InlineData("$A$1")]
        [InlineData("C$5")]
        [InlineData("$XFD1")]
        [InlineData("A$1048575")]
        [InlineData("XFD1048575")]
        public void SetCurrentCellAddressTest2(string address)
        {
            Worksheet worksheet = new Worksheet();
            Assert.Equal(0, worksheet.GetCurrentColumnNumber());
            Assert.Equal(0, worksheet.GetCurrentRowNumber());
            worksheet.GoToNextRow();
            worksheet.GoToNextColumn();
            worksheet.SetCurrentCellAddress(address);
            Address addr = new Address(address);
            Assert.Equal(addr.Column, worksheet.GetCurrentColumnNumber());
            Assert.Equal(addr.Row, worksheet.GetCurrentRowNumber());
        }

        [Theory(DisplayName = "Test of the failing SetCurrentCellAddress function on invalid columns or rows")]
        [InlineData(-1, 0)]
        [InlineData(0, -1)]
        [InlineData(-10, -10)]
        [InlineData(16384, 1048575)]
        [InlineData(16383, 1048576)]
        public void SetCurrentCellAddressFailTest(int column, int row)
        {
            Worksheet worksheet = new Worksheet();
            Assert.Throws<RangeException>(() => worksheet.SetCurrentCellAddress(column, row));
        }

        [Theory(DisplayName = "Test of the failing SetCurrentCellAddress function on an invalid address as string")]
        [InlineData(null)]
        [InlineData("")]
        [InlineData(":")]
        [InlineData("XFE1")]
        [InlineData("A1:A1")]
        [InlineData("A0")]
        [InlineData("A1048577")]
        public void SetCurrentCellAddressFailTest2(string address)
        {
            Worksheet worksheet = new Worksheet();
            Assert.ThrowsAny<Exception>(() => worksheet.SetCurrentCellAddress(address));
        }


        [Theory(DisplayName = "Test of the SetSelectedCells function with range objects")]
        [InlineData("A1:A1")]
        [InlineData("B2:C10")]
        [InlineData("A1:A10")]
        [InlineData("A1:R1")]
        [InlineData("$A$1:$R$1")]
        [InlineData("A1:XFD1048575")]
        public void SetSelectedCellsTest(string addressExpression)
        {
            Worksheet worksheet = new Worksheet();
            Assert.Empty(worksheet.SelectedCells);
            Range range = new Range(addressExpression);
            worksheet.AddSelectedCells(range);
            Assert.Contains(range, worksheet.SelectedCells);
        }

        [Theory(DisplayName = "Test of the SetSelectedCells function with strings")]
        [InlineData("A1:A1")]
        [InlineData("B2:C10")]
        [InlineData("C10:B5")]
        [InlineData("A1:A10")]
        [InlineData("A1:R1")]
        [InlineData("$A$1:$R$1")]
        [InlineData("A1:XFD1048575")]
        [InlineData(null)]
        public void SetSelectedCellsTest2(string addressExpression)
        {
            Worksheet worksheet = new Worksheet();
            Assert.Empty(worksheet.SelectedCells);
            worksheet.AddSelectedCells(addressExpression);
            if (addressExpression == null)
            {
                Assert.Empty(worksheet.SelectedCells);
            }
            else
            {
                Range range = new Range(addressExpression);
                Assert.Contains(range, worksheet.SelectedCells);
            }
        }


        [Theory(DisplayName = "Test of the SetSelectedCells function with address objects")]
        [InlineData("A1", "A1")]
        [InlineData("B2", "C10")]
        [InlineData("C10", "B5")]
        [InlineData("A1", "A10")]
        [InlineData("A1", "R1")]
        [InlineData("$A$1", "$R$1")]
        [InlineData("A1", "XFD1048575")]
        public void SetSelectedCellsTest3(string startAddress, string endAddress)
        {
            Worksheet worksheet = new Worksheet();
            Assert.Empty(worksheet.SelectedCells);
            Address start = new Address(startAddress);
            Address end = new Address(endAddress);
            Range range = new Range(start, end);
            worksheet.AddSelectedCells(start, end);
            Assert.Contains(range, worksheet.SelectedCells);
        }

        [Theory(DisplayName = "Test of the SetSheetProtectionPassword function")]
        [InlineData(null, null, false)]
        [InlineData("", null, false)]
        [InlineData("x", "x", true)]
        [InlineData("***", "***", true)]
        public void SetSheetProtectionPasswordTest(string password, string expectedPassword, bool expectedUsage)
        {
            Worksheet worksheet = new Worksheet();
            Assert.False(worksheet.UseSheetProtection);
            Assert.Null(worksheet.SheetProtectionPassword);
            worksheet.SetSheetProtectionPassword(password);
            Assert.Equal(expectedUsage, worksheet.UseSheetProtection);
            Assert.Equal(expectedPassword, worksheet.SheetProtectionPassword);
        }

        [Theory(DisplayName = "Test of the SetSheetname function")]
        [InlineData("1", true, "1")]
        [InlineData("test", true, "test")]
        [InlineData("test-test", true, "test-test")]
        [InlineData("$$$", true, "$$$")]
        [InlineData("a b", true, "a b")]
        [InlineData("a\tb", true, "a\tb")]
        [InlineData("-------------------------------", true, "-------------------------------")]
        [InlineData("", false, null)]
        [InlineData(null, false, null)]
        [InlineData("a[b", false, null)]
        [InlineData("a]b", false, null)]
        [InlineData("a*b", false, null)]
        [InlineData("a?b", false, null)]
        [InlineData("a/b", false, null)]
        [InlineData("a\\b", false, null)]
        [InlineData("--------------------------------", false, null)]
        public void SetSheetnameTest(string name, bool expectedValid, string expectedName)
        {
            Worksheet worksheet = new Worksheet();
            Assert.Null(worksheet.SheetName);
            if (expectedValid)
            {
                worksheet.SetSheetName(name);
                Assert.Equal(expectedName, worksheet.SheetName);
            }
            else
            {
                Assert.Throws<FormatException>(() => worksheet.SetSheetName(name));
            }
        }

        [Theory(DisplayName = "Test of the SetSheetName function with sanitation when one worksheet already exists")]
        [InlineData(false, "test", true, "test")]
        [InlineData(false, "", false, null)]
        [InlineData(false, null, false, null)]
        [InlineData(false, "a[b", false, null)]
        [InlineData(false, "a]b", false, null)]
        [InlineData(false, "a*b", false, null)]
        [InlineData(false, "a?b", false, null)]
        [InlineData(false, "a/b", false, null)]
        [InlineData(false, "a\\b", false, null)]
        [InlineData(false, "--------------------------------", false, null)]
        [InlineData(true, "test", true, "test")]
        [InlineData(true, "", true, "Sheet2")]
        [InlineData(true, null, true, "Sheet2")]
        [InlineData(true, "a[b", true, "a_b")]
        [InlineData(true, "a]b", true, "a_b")]
        [InlineData(true, "a*b", true, "a_b")]
        [InlineData(true, "a?b", true, "a_b")]
        [InlineData(true, "a/b", true, "a_b")]
        [InlineData(true, "a\\b", true, "a_b")]
        [InlineData(true, "--------------------------------", true, "-------------------------------")]
        public void SetSheetNameTest2(bool useSanitation, string name, bool expectedValid, string expectedName)
        {
            Workbook workbook = new Workbook("Sheet1");
            workbook.AddWorksheet("test");
            Worksheet worksheet = workbook.CurrentWorksheet;
            Assert.Equal("test", worksheet.SheetName);
            if (expectedValid)
            {
                worksheet.SetSheetName(name, useSanitation);
                Assert.Equal(expectedName, worksheet.SheetName);
            }
            else
            {
                Assert.Throws<FormatException>(() => worksheet.SetSheetName(name, useSanitation));
            }
        }

        [Fact(DisplayName = "Test of the failing SetSheetName function with sanitizing on a missing Workbook reference")]
        public void SetSheetNameFailingTest()
        {
            Worksheet worksheet = new Worksheet(); // Worksheet was not created over a workbook
            Assert.Throws<WorksheetException>(() => worksheet.SetSheetName("test", true));
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
