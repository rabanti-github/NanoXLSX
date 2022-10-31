using NanoXLSX;
using NanoXLSX.Shared.Exceptions;
using NanoXLSX.Styles;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xunit;

namespace NanoXLSX_Test.Workbooks
{
    public class CopyWorksheetTest
    {

        [Theory(DisplayName = "Test of the 'CopyWorksheetIntoThis' function by name")]
        [InlineData(false, "worksheet1", "worksheet2", "copy", "copy")]
        [InlineData(true, "worksheet1", "worksheet2", "copy", "copy")]
        [InlineData(true, "worksheet1", "worksheet2", "worksheet1", "worksheet2")]
        public void CopyWorksheetIntoThisTest(bool sanitize, string givenWorksheetName1, string givenSourceWsName, string copyName, string expectedTargetWsName)
        {
            Workbook workbook1 = new Workbook(givenWorksheetName1);
            Worksheet worksheet2 = createWorksheet();
            worksheet2.SheetName = givenSourceWsName;
            workbook1.AddWorksheet(worksheet2);
            workbook1.CopyWorksheetIntoThis(givenSourceWsName, copyName, sanitize);
            AssertCopy(workbook1, givenSourceWsName, workbook1, expectedTargetWsName);
        }

        [Fact(DisplayName = "Test of the failing 'CopyWorksheetIntoThis' (name) function when with disabled sanitation on duplicate worksheets")]
        public void CopyWorksheetIntoThisFailTest()
        {
            Workbook workbook1 = new Workbook("worksheet1");
            Worksheet worksheet2 = createWorksheet();
            worksheet2.SheetName = "worksheet2";
            workbook1.AddWorksheet(worksheet2);
            Assert.ThrowsAny<WorksheetException>(() => workbook1.CopyWorksheetIntoThis("worksheet2", "worksheet1", false));
        }

        [Theory(DisplayName = "Test of the 'CopyWorksheetIntoThis' function by index")]
        [InlineData(false, "worksheet1", "worksheet2", "copy", "copy")]
        [InlineData(true, "worksheet1", "worksheet2", "copy", "copy")]
        [InlineData(true, "worksheet1", "worksheet2", "worksheet1", "worksheet2")]
        public void CopyWorksheetIntoThisTest2(bool sanitize, string givenWorksheetName1, string givenSourceWsName, string copyName, string expectedTargetWsName)
        {
            Workbook workbook1 = new Workbook(givenWorksheetName1);
            Worksheet worksheet2 = createWorksheet();
            worksheet2.SheetName = givenSourceWsName;
            workbook1.AddWorksheet(worksheet2);
            workbook1.CopyWorksheetIntoThis(1, copyName, sanitize);
            AssertCopy(workbook1, givenSourceWsName, workbook1, expectedTargetWsName);
        }

        [Fact(DisplayName = "Test of the failing 'CopyWorksheetIntoThis' (index) function when with disabled sanitation on duplicate worksheets")]
        public void CopyWorksheetIntoThisFailTest2()
        {
            Workbook workbook1 = new Workbook("worksheet1");
            Worksheet worksheet2 = createWorksheet();
            worksheet2.SheetName = "worksheet2";
            workbook1.AddWorksheet(worksheet2);
            Assert.ThrowsAny<WorksheetException>(() => workbook1.CopyWorksheetIntoThis(1, "worksheet1", false));
        }

        [Theory(DisplayName = "Test of the 'CopyWorksheetIntoThis' function by reference")]
        [InlineData(false, "worksheet1", "worksheet2", "copy", "copy")]
        [InlineData(true, "worksheet1", "worksheet2", "copy", "copy")]
        [InlineData(true, "worksheet1", "worksheet2", "worksheet1", "worksheet2")]
        public void CopyWorksheetIntoThisTest3(bool sanitize, string givenWorksheetName1, string givenSourceWsName, string copyName, string expectedTargetWsName)
        {
            Workbook workbook1 = new Workbook(givenWorksheetName1);
            Worksheet worksheet2 = createWorksheet();
            worksheet2.SheetName = givenSourceWsName;
            workbook1.AddWorksheet(worksheet2);
            workbook1.CopyWorksheetIntoThis(worksheet2, copyName, sanitize);
            AssertCopy(workbook1, givenSourceWsName, workbook1, expectedTargetWsName);
        }

        [Fact(DisplayName = "Test of the failing 'CopyWorksheetIntoThis' (reference) function when with disabled sanitation on duplicate worksheets")]
        public void CopyWorksheetIntoThisFailTest3()
        {
            Workbook workbook1 = new Workbook("worksheet1");
            Worksheet worksheet2 = createWorksheet();
            worksheet2.SheetName = "worksheet2";
            workbook1.AddWorksheet(worksheet2);
            Assert.ThrowsAny<WorksheetException>(() => workbook1.CopyWorksheetIntoThis(worksheet2, "worksheet1", false));
        }


        [Theory(DisplayName = "Test of the 'CopyWorksheetTo' function by name")]
        [InlineData(false, "worksheet1", "worksheet2", "copy", "copy")]
        [InlineData(true, "worksheet1", "worksheet2", "copy", "copy")]
        [InlineData(true, "worksheet1", "worksheet2", "worksheet1", "worksheet2")]
        public void CopyWorksheetToTest(bool sanitize, string givenWorksheetName1, string givenSourceWsName, string copyName, string expectedTargetWsName)
        {
            Workbook workbook1 = new Workbook(givenWorksheetName1);
            Workbook workbook2 = new Workbook(givenWorksheetName1);
            Worksheet worksheet2 = createWorksheet();
            worksheet2.SheetName = givenSourceWsName;
            workbook1.AddWorksheet(worksheet2);
            workbook1.CopyWorksheetTo(givenSourceWsName, copyName, workbook2, sanitize);
            AssertCopy(workbook1, givenSourceWsName, workbook2, expectedTargetWsName);
        }

        [Fact(DisplayName = "Test of the failing 'CopyWorksheetTo' (name) function when with disabled sanitation on duplicate worksheets")]
        public void CopyWorksheetToFailTest()
        {
            Workbook workbook1 = new Workbook("worksheet1");
            Workbook workbook2 = new Workbook("worksheet1");
            Worksheet worksheet2 = createWorksheet();
            worksheet2.SheetName = "worksheet2";
            workbook1.AddWorksheet(worksheet2);
            Assert.ThrowsAny<WorksheetException>(() => workbook1.CopyWorksheetTo("worksheet2", "worksheet1", workbook2, false));
        }

        [Theory(DisplayName = "Test of the 'CopyWorksheetTo' function by index")]
        [InlineData(false, "worksheet1", "worksheet2", "copy", "copy")]
        [InlineData(true, "worksheet1", "worksheet2", "copy", "copy")]
        [InlineData(true, "worksheet1", "worksheet2", "worksheet1", "worksheet2")]
        public void CopyWorksheetToTest2(bool sanitize, string givenWorksheetName1, string givenSourceWsName, string copyName, string expectedTargetWsName)
        {
            Workbook workbook1 = new Workbook(givenWorksheetName1);
            Workbook workbook2 = new Workbook(givenWorksheetName1);
            Worksheet worksheet2 = createWorksheet();
            worksheet2.SheetName = givenSourceWsName;
            workbook1.AddWorksheet(worksheet2);
            workbook1.CopyWorksheetTo(1, copyName, workbook2, sanitize);
            AssertCopy(workbook1, givenSourceWsName, workbook2, expectedTargetWsName);
        }

        [Fact(DisplayName = "Test of the failing 'CopyWorksheetTo' (index) function when with disabled sanitation on duplicate worksheets")]
        public void CopyWorksheetToFailTest2()
        {
            Workbook workbook1 = new Workbook("worksheet1");
            Workbook workbook2 = new Workbook("worksheet1");
            Worksheet worksheet2 = createWorksheet();
            worksheet2.SheetName = "worksheet2";
            workbook1.AddWorksheet(worksheet2);
            Assert.ThrowsAny<WorksheetException>(() => workbook1.CopyWorksheetTo(1, "worksheet1", workbook2, false));
        }

        [Theory(DisplayName = "Test of the 'CopyWorksheetTo' function by reference")]
        [InlineData(false, "worksheet1", "worksheet2", "copy", "copy")]
        [InlineData(true, "worksheet1", "worksheet2", "copy", "copy")]
        [InlineData(true, "worksheet1", "worksheet2", "worksheet1", "worksheet2")]
        public void CopyWorksheetToTest3(bool sanitize, string givenWorksheetName1, string givenSourceWsName, string copyName, string expectedTargetWsName)
        {
            Workbook workbook1 = new Workbook(givenWorksheetName1);
            Workbook workbook2 = new Workbook(givenWorksheetName1);
            Worksheet worksheet2 = createWorksheet();
            worksheet2.SheetName = givenSourceWsName;
            workbook1.AddWorksheet(worksheet2);
            Workbook.CopyWorksheetTo(worksheet2, copyName, workbook2, sanitize);
            AssertCopy(workbook1, givenSourceWsName, workbook2, expectedTargetWsName);
        }

        [Fact(DisplayName = "Test of the failing 'CopyWorksheetTo' (reference) function when with disabled sanitation on duplicate worksheets")]
        public void CopyWorksheetToFailTest4()
        {
            Workbook workbook1 = new Workbook("worksheet1");
            Workbook workbook2 = new Workbook("worksheet1");
            Worksheet worksheet2 = createWorksheet();
            worksheet2.SheetName = "worksheet2";
            workbook1.AddWorksheet(worksheet2);
            Assert.ThrowsAny<WorksheetException>(() => Workbook.CopyWorksheetTo(worksheet2, "worksheet1", workbook2, false));
        }

        [Fact(DisplayName = "Test of the failing 'CopyWorksheetTo' function when no Workbook was defined")]
        public void CopyWorksheetToFailTest5()
        {
            Workbook workbook1 = null;
            Worksheet worksheet2 = createWorksheet();
            Assert.ThrowsAny<WorksheetException>(() => Workbook.CopyWorksheetTo(worksheet2, "copy", workbook1));
        }

        [Fact(DisplayName = "Test of the failing 'CopyWorksheetTo' function when no worksheet was defined")]
        public void CopyWorksheetToFailTest6()
        {
            Workbook workbook1 = new Workbook("worksheet1");
            Worksheet worksheet2 = null;
            Assert.ThrowsAny<WorksheetException>(() => Workbook.CopyWorksheetTo(worksheet2, "copy", workbook1));
        }


        [Fact(DisplayName = "Test of the 'Copy' function within the Worksheet class")]
        public void CopyTest()
        {
            Worksheet worksheet = createWorksheet();
            Worksheet worksheet2 = worksheet.Copy();
            AssertWorksheetCopy(worksheet, worksheet2);
        }

        private void AssertCopy(Workbook sourceWorkbook, string sourceName, Workbook targetWorkbook, string targetName)
        {
            Worksheet w1 = sourceWorkbook.GetWorksheet(sourceName);
            Worksheet w2 = targetWorkbook.GetWorksheet(targetName);
            AssertWorksheetCopy(w1, w2);
        }

        [Fact(DisplayName = "Test of the 'CopyWorksheetTo' function for proper saving")]
        public void CopyWorksheetSaveTest()
        {
            Workbook workbook1 = new Workbook("worksheet1");
            Workbook workbook2 = new Workbook("worksheet1b");
            Worksheet worksheet2 = createWorksheet();
            worksheet2.SheetName = "worksheet2";
            workbook1.AddWorksheet(worksheet2);
            Workbook.CopyWorksheetTo(worksheet2, "copy", workbook2);

            Workbook newWorkbook = TestUtils.WriteAndReadWorkbook(workbook2);
            Assert.Equal(workbook2.Worksheets.Count, newWorkbook.Worksheets.Count);
        }

        private void AssertWorksheetCopy(Worksheet w1, Worksheet w2)
        {
            IEnumerable<string> keys = w1.Cells.Keys.AsEnumerable();
            Assert.Equal(w2.Cells.Count, keys.Count());
            foreach (string address in keys)
            {
                Cell c1 = w1.Cells[address];
                Cell c2 = w2.Cells[address];
                Assert.Equal(c2.CellAddress, c1.CellAddress);
                Assert.Equal(c2.Value, c1.Value);
                Assert.Equal(c2.CellAddressType, c1.CellAddressType);
                AssertStyle(c1.CellStyle, c2.CellStyle);
                Assert.Equal(c2.DataType, c1.DataType);
            }
            Assert.Equal(w2.ActivePane, w1.ActivePane);
            AssertStyle(w1.ActiveStyle, w2.ActiveStyle);
            Assert.Equal(w2.AutoFilterRange, w1.AutoFilterRange);
            // columns
            IEnumerable<int> keys2 = w1.Columns.Keys.AsEnumerable();
            Assert.Equal(w2.Columns.Count, keys2.Count());
            foreach (int col in keys2)
            {
                Column c1 = w1.Columns[col];
                Column c2 = w2.Columns[col];
                Assert.Equal(c2.ColumnAddress, c1.ColumnAddress);
                Assert.Equal(c2.HasAutoFilter, c1.HasAutoFilter);
                Assert.Equal(c2.IsHidden, c1.IsHidden);
                Assert.Equal(c2.Number, c1.Number);
                Assert.Equal(c2.Width, c1.Width);
            }
            Assert.Equal(w2.CurrentCellDirection, w1.CurrentCellDirection);
            Assert.Equal(w2.DefaultColumnWidth, w1.DefaultColumnWidth);
            Assert.Equal(w2.DefaultRowHeight, w1.DefaultRowHeight);
            Assert.Equal(w2.FreezeSplitPanes, w1.FreezeSplitPanes);
            Assert.Equal(w2.Hidden, w1.Hidden);
            keys2 = w1.HiddenRows.Keys.AsEnumerable();
            Assert.Equal(w2.HiddenRows.Count, keys2.Count());
            foreach (int row in keys2)
            {
                Assert.Equal(w2.HiddenRows[row], w1.HiddenRows[row]);
            }
            keys = w1.MergedCells.Keys.AsEnumerable();
            Assert.Equal(w2.MergedCells.Count, keys.Count());
            foreach (string address in keys)
            {
                NanoXLSX.Range r1 = w1.MergedCells[address];
                NanoXLSX.Range r2 = w2.MergedCells[address];
                Assert.Equal(r2.StartAddress, r1.StartAddress);
                Assert.Equal(r2.EndAddress, r1.EndAddress);
            }
            Assert.Equal(w2.PaneSplitAddress, w2.PaneSplitAddress);
            Assert.Equal(w2.PaneSplitLeftWidth, w2.PaneSplitLeftWidth);
            Assert.Equal(w2.PaneSplitTopHeight, w2.PaneSplitTopHeight);
            Assert.Equal(w2.PaneSplitTopLeftCell, w2.PaneSplitTopLeftCell);
            keys2 = w1.RowHeights.Keys.AsEnumerable();
            Assert.Equal(w2.RowHeights.Count, keys2.Count());
            foreach (int row in keys2)
            {
                Assert.Equal(w2.RowHeights[row], w1.RowHeights[row]);
            }
            Assert.Equal(w2.SelectedCells, w1.SelectedCells);
            Assert.Equal(w2.SheetProtectionPassword, w1.SheetProtectionPassword);
            Assert.Equal(w2.SheetProtectionPasswordHash, w1.SheetProtectionPasswordHash);
            Assert.Equal(w2.SheetProtectionValues.Count, w1.SheetProtectionValues.Count);
            for (int i = 0; i < w1.SheetProtectionValues.Count; i++)
            {
                Assert.Equal(w2.SheetProtectionValues[i], w1.SheetProtectionValues[i]);
            }
            Assert.Equal(w2.UseSheetProtection, w1.UseSheetProtection);
        }

        private void AssertStyle(Style style1, Style style2)
        {
            if (style1 == null)
            {
                Assert.Null(style2);
            }
            else
            {
                Assert.Equal(style2.GetHashCode(), style1.GetHashCode());
            }
        }

        private Worksheet createWorksheet()
        {
            Worksheet w = new Worksheet();
            Style s1 = BasicStyles.BoldItalic;
            Style s2 = BasicStyles.Bold.Append(BasicStyles.DateFormat);
            w.AddCell("A1", "A1", s1);
            w.AddCell(true, "B2");
            w.AddCell(100, "C3", s2);
            w.AddCell(2.23f, "D4");
            w.AddCell(false, "D5");
            w.AddCellFormula("=A2", "E5");
            w.SetColumnWidth(2, 31.2f);
            w.SetRowHeight(2, 50.6f);
            w.AddHiddenColumn(1);
            w.AddHiddenColumn(3);
            w.AddAllowedActionOnSheetProtection(Worksheet.SheetProtectionValue.sort);
            w.AddAllowedActionOnSheetProtection(Worksheet.SheetProtectionValue.autoFilter);
            w.SetSheetProtectionPassword("pwd");
            w.AddHiddenRow(1);
            w.AddHiddenRow(3);
            w.CurrentCellDirection = Worksheet.CellDirection.Disabled;
            w.DefaultColumnWidth = 55.5f;
            w.DefaultRowHeight = 45.3f;
            w.Hidden = true;
            w.MergeCells(new NanoXLSX.Range("D4:D5"));
            w.SetActiveStyle(s2);
            w.SetAutoFilter("B1:C2");
            w.SetCurrentCellAddress("D5");
            w.SetSelectedCells(new NanoXLSX.Range("C3:C3"));
            w.UseSheetProtection = true;
            w.SetSplit(3, 2, true, new Address("F4"), Worksheet.WorksheetPane.bottomRight);
            return w;
        }
    }
}
