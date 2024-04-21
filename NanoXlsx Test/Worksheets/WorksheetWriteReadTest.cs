using NanoXLSX;
using NanoXLSX.Styles;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xunit;
using Range = NanoXLSX.Range;

namespace NanoXLSX_Test.Worksheets
{
    public class WorksheetWriteReadTest
    {

        [Theory(DisplayName = "Test of the 'AutoFilterRange' property when writing and reading a worksheet")]
        [InlineData(null, 0)]
        [InlineData("A1:A1", 0)]
        [InlineData("A1:C1", 0)]
        [InlineData("B1:D1", 0)]
        [InlineData("A1", 0)]
        [InlineData(null, 1)]
        [InlineData("A1:A1", 1)]
        [InlineData("A1:C1", 2)]
        [InlineData("B1:D1", 3)]
        [InlineData("B1", 2)]
        public void AutoFilterRangeWriteReadTest(string autoFilterRange, int sheetIndex)
        {
            Workbook workbook = PrepareWorkbook(4, "test");
            Range? range = null;
            if (autoFilterRange != null)
            {
                range = new Range(autoFilterRange);
                for (int i = 0; i <= sheetIndex; i++)
                {
                    if (sheetIndex == i)
                    {
                        workbook.SetCurrentWorksheet(i);
                        workbook.CurrentWorksheet.SetAutoFilter(range.Value.StartAddress.Column, range.Value.EndAddress.Column);
                    }
                }
            }
            Worksheet givenWorksheet = WriteAndReadWorksheet(workbook, sheetIndex);
            if (autoFilterRange == null)
            {
                Assert.Null(givenWorksheet.AutoFilterRange);
            }
            else
            {
                Assert.Equal(range, givenWorksheet.AutoFilterRange.Value);
            }
        }

        [Theory(DisplayName = "Test of the 'Columns' property when writing and reading a worksheet")]
        [InlineData("", 0, true, false)]
        [InlineData("0", 0, true, false)]
        [InlineData("0,1,2", 0, true, false)]
        [InlineData("1,3,5", 0, true, false)]
        [InlineData("", 1, true, false)]
        [InlineData("0", 1, true, false)]
        [InlineData("0,1,2", 2, true, false)]
        [InlineData("1,3,5", 3, true, false)]
        [InlineData("", 0, false, true)]
        [InlineData("0", 0, false, true)]
        [InlineData("0,1,2", 0, false, true)]
        [InlineData("1,3,5", 0, false, true)]
        [InlineData("", 1, false, true)]
        [InlineData("0", 1, false, true)]
        [InlineData("0,1,2", 2, false, true)]
        [InlineData("1,3,5", 3, false, true)]
        public void ColumnsWriteReadTest(string columnDefinitions, int sheetIndex, bool setWidth, bool setHidden)
        {
            string[] tokens = columnDefinitions.Split(',');
            List<int> columnIndices = new List<int>();
            foreach (string token in tokens)
            {
                if (token != "")
                {
                    columnIndices.Add(int.Parse(token));
                }
            }
            Workbook workbook = PrepareWorkbook(4, "test");
            for (int i = 0; i <= sheetIndex; i++)
            {
                if (sheetIndex == i)
                {
                    workbook.SetCurrentWorksheet(i);
                    foreach(int index in columnIndices)
                    {
                        if (setWidth)
                        {
                            workbook.CurrentWorksheet.SetColumnWidth(index, 99);
                        }
                        if (setHidden)
                        {
                            workbook.CurrentWorksheet.AddHiddenColumn(index);
                        }
                    }
                }
            }
            Worksheet givenWorksheet = WriteAndReadWorksheet(workbook, sheetIndex);
            Assert.Equal(columnIndices.Count, givenWorksheet.Columns.Count);
            foreach(KeyValuePair<int,Column> column in givenWorksheet.Columns)
            {
                Assert.Contains(columnIndices, x => x == column.Value.Number);
                if (setWidth)
                {
                   
                    Assert.True(Math.Abs(column.Value.Width - Utils.GetInternalColumnWidth(99)) < 0.001);
                }
                if (setHidden)
                {
                    Assert.True(column.Value.IsHidden);
                }
            }
        }

        [Theory(DisplayName = "Test of the 'DefaultColumnWidth' property when writing and reading a worksheet")]
        [InlineData(1f, 0)]
        [InlineData(11f, 0)]
        [InlineData(55.55f, 0)]
        [InlineData(1f, 1)]
        [InlineData(11f, 2)]
        [InlineData(55.55f, 3)]
        public void DefaultColumnWidthWriteReadTest(float width, int sheetIndex)
        {
            Workbook workbook = PrepareWorkbook(4, "test");
            for (int i = 0; i <= sheetIndex; i++)
            {
                if (sheetIndex == i)
                {
                    workbook.SetCurrentWorksheet(i);
                    workbook.CurrentWorksheet.DefaultColumnWidth = width;
                }
            }
            Worksheet givenWorksheet = WriteAndReadWorksheet(workbook, sheetIndex);
            Assert.True(Math.Abs(givenWorksheet.DefaultColumnWidth - width) < 0.001);
        }

        [Theory(DisplayName = "Test of the 'DefaultRowHeight' property when writing and reading a worksheet")]
        [InlineData(1f, 0)]
        [InlineData(11f, 0)]
        [InlineData(55.55f, 0)]
        [InlineData(1f, 1)]
        [InlineData(11f, 2)]
        [InlineData(55.55f, 3)]
        public void DefaultRowHeightWriteReadTest(float height, int sheetIndex)
        {
            Workbook workbook = PrepareWorkbook(4, "test");
            for (int i = 0; i <= sheetIndex; i++)
            {
                if (sheetIndex == i)
                {
                    workbook.SetCurrentWorksheet(i);
                    workbook.CurrentWorksheet.DefaultRowHeight = height;
                }
            }
            Worksheet givenWorksheet = WriteAndReadWorksheet(workbook, sheetIndex);
            Assert.True(Math.Abs(givenWorksheet.DefaultRowHeight - height) < 0.001);
        }

        [Theory(DisplayName = "Test of the 'HiddenRows' property when writing and reading a worksheet")]
        [InlineData("", 0)]
        [InlineData("0", 0)]
        [InlineData("0,1,2", 0)]
        [InlineData("1,3,5", 0)]
        [InlineData("", 1)]
        [InlineData("0", 1)]
        [InlineData("0,1,2", 2)]
        [InlineData("1,3,5", 3)]
        public void HiddenRowsWriteReadTest(string rowDefinitions, int sheetIndex)
        {
            string[] tokens = rowDefinitions.Split(',');
            List<int> rowIndices = new List<int>();
            foreach (string token in tokens)
            {
                if (token != "")
                {
                    rowIndices.Add(int.Parse(token));
                }
            }
            Workbook workbook = PrepareWorkbook(4, "test");
            for (int i = 0; i <= sheetIndex; i++)
            {
                if (sheetIndex == i)
                {
                    workbook.SetCurrentWorksheet(i);
                    foreach (int index in rowIndices)
                    {
                        workbook.CurrentWorksheet.AddHiddenRow(index);
                    }
                }
            }
            Worksheet givenWorksheet = WriteAndReadWorksheet(workbook, sheetIndex);
            Assert.Equal(rowIndices.Count, givenWorksheet.HiddenRows.Count);
            foreach (KeyValuePair<int, bool> hiddenRow in givenWorksheet.HiddenRows)
            {
                Assert.Contains(rowIndices, x => x == hiddenRow.Key);
                Assert.True(hiddenRow.Value);
            }
        }



        [Theory(DisplayName = "Test of the 'RowHeight' property when writing and reading a worksheet")]
        [InlineData("", "", 0)]
        [InlineData("0", "17", 0)]
        [InlineData("0,1,2", "11,12,13.5", 0)]
        [InlineData("1,3,5", "55.5,1.111,5.587", 0)]
        [InlineData("", "", 1)]
        [InlineData("0", "17.2", 1)]
        [InlineData("0,1,2", "11.05,12.1,13.55", 2)]
        [InlineData("1,3,5", "55.5,1.111,5.587", 3)]
        public void RowHeightsWriteReadTest(string rowDefinitions, string heightDefinitions, int sheetIndex)
        {
            string[] tokens = rowDefinitions.Split(',');
            string[] heightTokens = heightDefinitions.Split(',');
            Dictionary<int, float> rows = new Dictionary<int,float>();
            for (int i = 0; i < tokens.Length; i++)
            {
                if (tokens[i] != "")
                {
                    rows.Add(int.Parse(tokens[i]), float.Parse(heightTokens[i]));
                }
            }
            Workbook workbook = PrepareWorkbook(4, "test");
            for (int i = 0; i <= sheetIndex; i++)
            {
                if (sheetIndex == i)
                {
                    workbook.SetCurrentWorksheet(i);
                    foreach (KeyValuePair<int, float> row in rows)
                    {
                        workbook.CurrentWorksheet.SetRowHeight(row.Key, row.Value);
                    }
                }
            }
            Worksheet givenWorksheet = WriteAndReadWorksheet(workbook, sheetIndex);
            Assert.Equal(rows.Count, givenWorksheet.RowHeights.Count);
            foreach (KeyValuePair<int, float> rowHeight in givenWorksheet.RowHeights)
            {
                Assert.Contains(rows.Keys, x => x == rowHeight.Key);
                float expectedHeight = Utils.GetInternalRowHeight(rows[rowHeight.Key]);
                Assert.Equal(expectedHeight, rowHeight.Value);
            }
        }

        [Fact(DisplayName = "Test of the 'RowHeight' property when writing and reading a worksheet, if a row already exists")]
        public void RowHeightsWriteReadTest2()
        {
            Workbook workbook = new Workbook("worksheet1");
            workbook.CurrentWorksheet.AddCell(42, "C2");
            workbook.CurrentWorksheet.SetRowHeight(2, 22.55f);
            workbook.CurrentWorksheet.AddHiddenRow(2);
            Worksheet givenWorksheet = WriteAndReadWorksheet(workbook, 0);
            Assert.Equal(Utils.GetInternalRowHeight(22.55f), givenWorksheet.RowHeights[2]);
            Assert.True(givenWorksheet.HiddenRows[2]);
    }


        [Theory(DisplayName = "Test of the 'MergedCells' property when writing and reading a worksheet")]
        [InlineData(null, 0)]
        [InlineData("A1:A1", 0)]
        [InlineData("A1:C1", 0)]
        [InlineData("B1:D1", 0)]
        [InlineData("B1:D1,E5:E7", 0)]
        [InlineData(null, 1)]
        [InlineData("A1:A1", 1)]
        [InlineData("A1:C1", 2)]
        [InlineData("B1:D1", 3)]
        [InlineData("B1:D1,E5:E7", 3)]
        public void MergedCellsWriteReadTest(string mergedCellsRanges, int sheetIndex)
        {
            Workbook workbook = PrepareWorkbook(4, "test");
            List<Range>ranges = new List<Range>();
            if (mergedCellsRanges != null)
            {
                string[] split = mergedCellsRanges.Split(",");
                foreach(string range in split)
                {
                    ranges.Add(new Range(range));
                }
                for (int i = 0; i <= sheetIndex; i++)
                {
                    if (sheetIndex == i)
                    {
                        workbook.SetCurrentWorksheet(i);
                        foreach(Range range in ranges)
                        {
                            workbook.CurrentWorksheet.MergeCells(range);
                        }
                    }
                }
            }
            Worksheet givenWorksheet = WriteAndReadWorksheet(workbook, sheetIndex);
            if (mergedCellsRanges == null)
            {
                Assert.Empty(givenWorksheet.MergedCells);
            }
            else
            {
                foreach(Range range in ranges)
                {
                    Assert.Equal(range, givenWorksheet.MergedCells[range.ToString()]);
                }
            }
        }

        [Theory(DisplayName = "Test of the 'SelectedCells' property when writing and reading a worksheet")]
        [InlineData(null, 0)]
        [InlineData("A1:A1", 0)]
        [InlineData("A1:C1", 0)]
        [InlineData("B1:D1", 0)]
        [InlineData(null, 1)]
        [InlineData("A1:A1", 1)]
        [InlineData("A1:C1", 2)]
        [InlineData("B1:D1", 3)]
        public void SelectedCellsWriteReadTest(string selectedCellsRange, int sheetIndex)
        {
            Workbook workbook = PrepareWorkbook(4, "test");
            Range? range = null;
            if (selectedCellsRange != null)
            {
                range = new Range(selectedCellsRange);
                for (int i = 0; i <= sheetIndex; i++)
                {
                    if (sheetIndex == i)
                    {
                        workbook.SetCurrentWorksheet(i);
                        workbook.CurrentWorksheet.SetSelectedCells(range.Value);
                    }
                }
            }
            Worksheet givenWorksheet = WriteAndReadWorksheet(workbook, sheetIndex);
            if (selectedCellsRange == null)
            {
                Assert.Null(givenWorksheet.SelectedCells);
            }
            else
            {
                Assert.Equal(range.Value, givenWorksheet.SelectedCells);
            }
        }

        [Theory(DisplayName = "Test of the 'SelectedCellRanges' property when writing and reading a worksheet")]
        [InlineData(null, 0)]
        [InlineData("A1:A1", 0)]
        [InlineData("A1:C1", 0)]
        [InlineData("B1:D1", 0)]
        [InlineData(null, 1)]
        [InlineData("A1:A1", 1)]
        [InlineData("A1:C1", 2)]
        [InlineData("B1:D1", 3)]
        [InlineData("A1:A1,B1:B1", 0)]
        [InlineData("A1:C1,D1:F2", 0)]
        [InlineData("B1:D1,A1:A1,F3:F4", 0)]
        [InlineData("A1:A1,B1:B1", 1)]
        [InlineData("A1:C1,D1:F2", 2)]
        [InlineData("B1:D1,A1:A1,F3:F4", 3)]
        public void SelectedCellRangesWriteReadTest(string selectedCellsRanges, int sheetIndex)
        {
            Workbook workbook = PrepareWorkbook(4, "test");
            string[] ranges = null;
            if (selectedCellsRanges != null)
            {
                ranges = selectedCellsRanges.Split(',');
                foreach(string range in ranges)
                {
                    Range range2 = new Range(range);
                    for (int i = 0; i <= sheetIndex; i++)
                    {
                        if (sheetIndex == i)
                        {
                            workbook.SetCurrentWorksheet(i);
                            workbook.CurrentWorksheet.AddSelectedCells(range2);
                        }
                    }
                }
            }
            Worksheet givenWorksheet = WriteAndReadWorksheet(workbook, sheetIndex);
            if (selectedCellsRanges == null)
            {
                Assert.Empty(givenWorksheet.SelectedCellRanges);
            }
            else
            {
                Assert.Equal(ranges.Length, givenWorksheet.SelectedCellRanges.Count);
                foreach(string range in ranges)
                {
                    Assert.Contains(new Range(range), givenWorksheet.SelectedCellRanges);
                }
            }

        }

            [Fact(DisplayName = "Test of the 'SheetID'  property when writing and reading a worksheet")]
        public void SheetIDWriteReadTest()
        {
            Workbook workbook = new Workbook();
            string sheetName1 = "sheet_a";
            string sheetName2 = "sheet_b";
            string sheetName3 = "sheet_c";
            string sheetName4 = "sheet_d";
            int id1, id2, id3, id4;
            workbook.AddWorksheet(sheetName1);
            id1 = workbook.CurrentWorksheet.SheetID;
            workbook.AddWorksheet(sheetName2);
            id2 = workbook.CurrentWorksheet.SheetID;
            workbook.AddWorksheet(sheetName3);
            id3 = workbook.CurrentWorksheet.SheetID;
            workbook.AddWorksheet(sheetName4);
            id4 = workbook.CurrentWorksheet.SheetID;
            Workbook givenWorkbook = null;
            using (MemoryStream stream = new MemoryStream())
            {
                workbook.SaveAsStream(stream, true);
                stream.Position = 0;
                givenWorkbook = Workbook.Load(stream);
            }
            Assert.Equal(id1, givenWorkbook.Worksheets.First(w => w.SheetName == sheetName1).SheetID);
            Assert.Equal(id2, givenWorkbook.Worksheets.First(w => w.SheetName == sheetName2).SheetID);
            Assert.Equal(id3, givenWorkbook.Worksheets.First(w => w.SheetName == sheetName3).SheetID);
            Assert.Equal(id4, givenWorkbook.Worksheets.First(w => w.SheetName == sheetName4).SheetID);
        }

        [Fact(DisplayName = "Test of the 'SheetName'  property when writing and reading a worksheet")]
        public void SheetNameWriteReadTest()
        {
            Workbook workbook = new Workbook();
            string sheetName1 = "sheet_a";
            string sheetName2 = "sheet_b";
            string sheetName3 = "sheet_c";
            string sheetName4 = "sheet_d";
            int id1, id2, id3, id4;
            workbook.AddWorksheet(sheetName1);
            id1 = workbook.CurrentWorksheet.SheetID;
            workbook.AddWorksheet(sheetName2);
            id2 = workbook.CurrentWorksheet.SheetID;
            workbook.AddWorksheet(sheetName3);
            id3 = workbook.CurrentWorksheet.SheetID;
            workbook.AddWorksheet(sheetName4);
            id4 = workbook.CurrentWorksheet.SheetID;
            Workbook givenWorkbook = null;
            using (MemoryStream stream = new MemoryStream())
            {
                workbook.SaveAsStream(stream, true);
                stream.Position = 0;
                givenWorkbook = Workbook.Load(stream);
            }
            Assert.Equal(sheetName1, givenWorkbook.Worksheets.First(w => w.SheetID == id1).SheetName);
            Assert.Equal(sheetName3, givenWorkbook.Worksheets.First(w => w.SheetID == id3).SheetName);
            Assert.Equal(sheetName4, givenWorkbook.Worksheets.First(w => w.SheetID == id4).SheetName);
            Assert.Equal(sheetName2, givenWorkbook.Worksheets.First(w => w.SheetID == id2).SheetName);
        }

        [Theory(DisplayName = "Test of the 'SheetProtectionValues'  and 'UseSheetProtection' property when writing and reading a worksheet")]
        [InlineData(false, "", "", 0)]
        [InlineData(false, "autoFilter:0,sort:0", "", 0)]
        [InlineData(true, "", "objects:1,scenarios:1,selectLockedCells:1,selectUnlockedCells:1", 0)]
        [InlineData(true, "autoFilter:0", "autoFilter:0,objects:1,scenarios:1,selectLockedCells:1,selectUnlockedCells:1", 0)]
        [InlineData(true, "pivotTables:0", "pivotTables:0,objects:1,scenarios:1,selectLockedCells:1,selectUnlockedCells:1", 0)]
        [InlineData(true, "sort:0", "sort:0,objects:1,scenarios:1,selectLockedCells:1,selectUnlockedCells:1", 0)]
        [InlineData(true, "deleteRows:0", "deleteRows:0,objects:1,scenarios:1,selectLockedCells:1,selectUnlockedCells:1", 0)]
        [InlineData(true, "deleteColumns:0", "deleteColumns:0,objects:1,scenarios:1,selectLockedCells:1,selectUnlockedCells:1", 0)]
        [InlineData(true, "insertHyperlinks:0", "insertHyperlinks:0,objects:1,scenarios:1,selectLockedCells:1,selectUnlockedCells:1", 0)]
        [InlineData(true, "insertRows:0", "insertRows:0,objects:1,scenarios:1,selectLockedCells:1,selectUnlockedCells:1", 0)]
        [InlineData(true, "insertColumns:0", "insertColumns:0,objects:1,scenarios:1,selectLockedCells:1,selectUnlockedCells:1", 0)]
        [InlineData(true, "formatRows:0", "formatRows:0,objects:1,scenarios:1,selectLockedCells:1,selectUnlockedCells:1", 0)]
        [InlineData(true, "formatColumns:0", "formatColumns:0,objects:1,scenarios:1,selectLockedCells:1,selectUnlockedCells:1", 0)]
        [InlineData(true, "formatCells:0", "formatCells:0,objects:1,scenarios:1,selectLockedCells:1,selectUnlockedCells:1", 0)]
        [InlineData(true, "objects:0", "scenarios:1,selectLockedCells:1,selectUnlockedCells:1", 0)] 
        [InlineData(true, "scenarios:0", "objects:1,selectLockedCells:1,selectUnlockedCells:1", 0)]
        [InlineData(true, "selectLockedCells:0", "objects:1,scenarios:1", 0)]
        [InlineData(true, "selectUnlockedCells:0", "objects:1,scenarios:1,selectLockedCells:1,selectUnlockedCells:0", 0)]
        [InlineData(false, "", "", 1)]
        [InlineData(false, "autoFilter:0", "", 2)]
        [InlineData(true, "", "objects:1,scenarios:1,selectLockedCells:1,selectUnlockedCells:1", 3)]
        [InlineData(true, "autoFilter:0", "autoFilter:0,objects:1,scenarios:1,selectLockedCells:1,selectUnlockedCells:1", 1)]
        [InlineData(true, "pivotTables:0,sort:0", "pivotTables:0,sort:0,objects:1,scenarios:1,selectLockedCells:1,selectUnlockedCells:1", 2)]
        [InlineData(true, "sort:0,deleteColumns:0,formatCells:0", "sort:0,deleteColumns:0,formatCells:0,objects:1,scenarios:1,selectLockedCells:1,selectUnlockedCells:1", 3)]
        [InlineData(true, "deleteRows:0,formatCells:0", "deleteRows:0,formatCells:0,objects:1,scenarios:1,selectLockedCells:1,selectUnlockedCells:1", 1)]
        [InlineData(true, "deleteColumns:0,formatColumns:0,formatRows:0", "deleteColumns:0,formatColumns:0,formatRows:0,objects:1,scenarios:1,selectLockedCells:1,selectUnlockedCells:1", 2)]
        [InlineData(true, "insertHyperlinks:0,formatCells:0", "insertHyperlinks:0,formatCells:0,objects:1,scenarios:1,selectLockedCells:1,selectUnlockedCells:1", 3)]
        [InlineData(true, "insertRows:0,formatRows:0", "insertRows:0,formatRows:0,objects:1,scenarios:1,selectLockedCells:1,selectUnlockedCells:1", 1)]
        [InlineData(true, "insertColumns:0,formatColumns:0", "insertColumns:0,formatColumns:0,objects:1,scenarios:1,selectLockedCells:1,selectUnlockedCells:1", 2)]
        [InlineData(true, "formatRows:0,formatColumns:0", "formatRows:0,formatColumns:0,objects:1,scenarios:1,selectLockedCells:1,selectUnlockedCells:1", 3)]
        [InlineData(true, "formatColumns:0,formatCells:0", "formatColumns:0,formatCells:0,objects:1,scenarios:1,selectLockedCells:1,selectUnlockedCells:1", 1)]
        public void SheetProtectionWriteReadTest(bool useSheetProtection, string givenProtectionValues, string expectedProtectionValues, int sheetIndex)
        {
            Dictionary<Worksheet.SheetProtectionValue, bool> expectedProtection = PrepareSheetProtectionValues(expectedProtectionValues);
            Dictionary<Worksheet.SheetProtectionValue, bool> givenProtection = PrepareSheetProtectionValues(givenProtectionValues);
            Workbook workbook = PrepareWorkbook(4, "test");
            for (int i = 0; i <= sheetIndex; i++)
            {
                if (sheetIndex == i)
                {
                    workbook.SetCurrentWorksheet(i);
                    foreach (KeyValuePair<Worksheet.SheetProtectionValue, bool> item in givenProtection)
                    {
                       workbook.CurrentWorksheet.AddAllowedActionOnSheetProtection(item.Key);
                    }
                    // adding values will enable sheet protection in any case, can be deactivated afterwards
                    workbook.CurrentWorksheet.UseSheetProtection = useSheetProtection;
                }
            }
            Worksheet givenWorksheet = WriteAndReadWorksheet(workbook, sheetIndex);
            Assert.Equal(expectedProtection.Count, givenWorksheet.SheetProtectionValues.Count);
            Assert.Equal(useSheetProtection, givenWorksheet.UseSheetProtection);
            foreach (KeyValuePair<Worksheet.SheetProtectionValue, bool> item in expectedProtection)
            {
                if (item.Value)
                {
                    Assert.Contains(item.Key, givenWorksheet.SheetProtectionValues);
                }
            }
        }

        [Theory(DisplayName = "Test of the 'SheetProtectionPasswordHash' property when writing and reading a worksheet")]
        [InlineData("x", 0)]
        [InlineData("@test-1,23", 0)]
        [InlineData("", 0)]
        [InlineData(null, 0)]
        [InlineData("x", 1)]
        [InlineData("@test-1,23", 2)]
        [InlineData("", 3)]
        [InlineData(null, 4)]
        public void SheetProtectionPasswordHashWriteReadTest(string givenPassword, int sheetIndex)
        {
            string hash = null;
            Workbook workbook = PrepareWorkbook(5, "test");
            for (int i = 0; i <= sheetIndex; i++)
            {
                if (sheetIndex == i)
                {
                    workbook.SetCurrentWorksheet(i);
                    workbook.CurrentWorksheet.AddAllowedActionOnSheetProtection(Worksheet.SheetProtectionValue.deleteRows);
                    workbook.CurrentWorksheet.SetSheetProtectionPassword(givenPassword);
                    hash = workbook.CurrentWorksheet.SheetProtectionPasswordHash;
                }
            }
            Worksheet givenWorksheet = WriteAndReadWorksheet(workbook, sheetIndex);
            Assert.Equal(hash, givenWorksheet.SheetProtectionPasswordHash);
        }

        [Theory(DisplayName = "Test of the 'Hidden' property when writing and reading a worksheet")]
        [InlineData(false, 0)]
        [InlineData(true, 0)]
        [InlineData(false, 1)]
        [InlineData(true, 1)]
        [InlineData(false, 2)]
        [InlineData(true, 2)]
        public void HiddenWriteReadTest(bool hidden, int sheetIndex)
        {
            Workbook workbook = PrepareWorkbook(4, "test");
            for (int i = 0; i <= sheetIndex; i++)
            {
                if (i == 0 && i == sheetIndex)
                {
                    // Prevents setting selected worksheet as hidden
                    workbook.SetSelectedWorksheet(1);
                }
                if (sheetIndex == i)
                {
                    workbook.SetCurrentWorksheet(i);
                    workbook.CurrentWorksheet.Hidden = hidden;
                }
            }
            Worksheet givenWorksheet = WriteAndReadWorksheet(workbook, sheetIndex);
            Assert.Equal(hidden, givenWorksheet.Hidden);
        }

        [Fact(DisplayName = "Test of the replacement of the CellXF style part, applied on merged cells, when saving a workbook")]
        public void MixedCellsCellXfTest()
        {
            Workbook workbook = new Workbook("Sheet1");
            Style xfStyle = new Style();
            xfStyle.CurrentCellXf.Alignment = CellXf.TextBreakValue.shrinkToFit;
            Style style = BasicStyles.Bold.Append(xfStyle);
            workbook.CurrentWorksheet.AddCell("", "A1", style);
            workbook.CurrentWorksheet.AddCell("B", "A2", style);
            workbook.CurrentWorksheet.AddCell("", "A3", style);
            workbook.CurrentWorksheet.MergeCells(new Range("A1:A3"));
            Assert.False(workbook.CurrentWorksheet.Cells["A1"].CellStyle.CurrentCellXf.ForceApplyAlignment);
            Assert.False(workbook.CurrentWorksheet.Cells["A2"].CellStyle.CurrentCellXf.ForceApplyAlignment);
            Assert.False(workbook.CurrentWorksheet.Cells["A3"].CellStyle.CurrentCellXf.ForceApplyAlignment);
            Worksheet givenWorksheet = WriteAndReadWorksheet(workbook, 0);
            Assert.True(givenWorksheet.Cells["A1"].CellStyle.CurrentCellXf.ForceApplyAlignment);
            Assert.Equal(CellXf.TextBreakValue.shrinkToFit, givenWorksheet.Cells["A1"].CellStyle.CurrentCellXf.Alignment);
            Assert.True(givenWorksheet.Cells["A2"].CellStyle.CurrentCellXf.ForceApplyAlignment);
            Assert.Equal(CellXf.TextBreakValue.shrinkToFit, givenWorksheet.Cells["A2"].CellStyle.CurrentCellXf.Alignment);
            Assert.True(givenWorksheet.Cells["A3"].CellStyle.CurrentCellXf.ForceApplyAlignment);
            Assert.Equal(CellXf.TextBreakValue.shrinkToFit, givenWorksheet.Cells["A3"].CellStyle.CurrentCellXf.Alignment);
        }


        private static Dictionary<Worksheet.SheetProtectionValue, bool> PrepareSheetProtectionValues(string tokenString)
        {
            Dictionary<Worksheet.SheetProtectionValue, bool> dictionary = new Dictionary<Worksheet.SheetProtectionValue, bool>();
            string[] tokens = tokenString.Split(",");
            foreach (string token in tokens)
            {
                if (token == "")
                {
                    continue;
                }
                string[] subTokens = token.Split(":");
                Worksheet.SheetProtectionValue value = (Worksheet.SheetProtectionValue)Enum.Parse(typeof(Worksheet.SheetProtectionValue), subTokens[0]);
                if (subTokens[1] == "1")
                {
                    dictionary.Add(value, true);
                }
                else
                {
                    dictionary.Add(value, false);
                }
            }
            return dictionary;
        }

        private static Workbook PrepareWorkbook(int numberOfWorksheets, object a1Data)
        {
            Workbook workbook = new Workbook();
            for (int i = 0; i < numberOfWorksheets; i++)
            {
                workbook.AddWorksheet("worksheet" + (i + 1).ToString());
                workbook.CurrentWorksheet.AddCell(a1Data, "A1");
            }
            return workbook;
        }

        private static Worksheet WriteAndReadWorksheet(Workbook workbook, int worksheetIndex)
        {
            using (MemoryStream stream = new MemoryStream())
            {
                workbook.SaveAsStream(stream, true);
                stream.Position = 0;
                Workbook readWorkbook = Workbook.Load(stream);
                return readWorkbook.Worksheets[worksheetIndex];
            }
        }

    }
}
