using NanoXLSX;
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
        [InlineData(null, 1)]
        [InlineData("A1:A1", 1)]
        [InlineData("A1:C1", 2)]
        [InlineData("B1:D1", 3)]
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
