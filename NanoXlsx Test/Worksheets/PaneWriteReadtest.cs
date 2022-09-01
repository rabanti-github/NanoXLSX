using NanoXLSX;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xunit;

namespace NanoXLSX_Test.Worksheets
{
    public class PaneWriteReadtest
    {

        [Theory(DisplayName = "Test of the 'PaneSplitTopHeight' property when writing and reading a worksheet")]
        [InlineData(27f, null, 0)]
        [InlineData(100f, null, 0)]
        [InlineData(0f, null, 0)]
        [InlineData(27f, Worksheet.WorksheetPane.topLeft, 0)]
        [InlineData(100f, Worksheet.WorksheetPane.bottomLeft, 0)]
        [InlineData(0f, Worksheet.WorksheetPane.topRight, 0)]
        public void PaneSplitTopHeightWriteReadTest(float height, Worksheet.WorksheetPane? activePane, int sheetIndex)
        {
            Workbook workbook = PrepareWorkbook(4, "test");
            for (int i = 0; i <= sheetIndex; i++)
            {
                if (sheetIndex == i)
                {
                    workbook.SetCurrentWorksheet(i);
                    workbook.CurrentWorksheet.SetHorizontalSplit(height, new Address("A2"), activePane);
                }
            }
            Worksheet givenWorksheet = WriteAndReadWorksheet(workbook, sheetIndex);
            Assert.Equal(height, givenWorksheet.PaneSplitTopHeight);
        }

        [Theory(DisplayName = "Test of the 'PaneSplitTopHeight' property defined by a split address, when writing and reading a worksheet")]
        [InlineData(0, false, "A2", null, 0)]
        [InlineData(1, false, "A2", null, 0)]
        [InlineData(15, false, "A18", null, 0)]
        [InlineData(0, false, "A2", Worksheet.WorksheetPane.topLeft, 0)]
        [InlineData(1, false, "A2", Worksheet.WorksheetPane.bottomLeft, 0)]
        [InlineData(15, false, "A18", Worksheet.WorksheetPane.topRight, 0)]
        [InlineData(0, true, "A2", null, 0)]
        [InlineData(1, true, "A2", null, 0)]
        [InlineData(15, true, "A18", null, 0)]
        [InlineData(0, true, "A2", Worksheet.WorksheetPane.topLeft, 0)]
        [InlineData(1, true, "A2", Worksheet.WorksheetPane.bottomLeft, 0)]
        [InlineData(15, true, "A18", Worksheet.WorksheetPane.topRight, 0)]
        public void PaneSplitTopHeightWriteReadTest2(int rowNumber, bool freeze, string topLeftCellAddress, Worksheet.WorksheetPane? activePane, int sheetIndex)
        {
            Workbook workbook = PrepareWorkbook(4, "test");
            for (int i = 0; i <= sheetIndex; i++)
            {
                if (sheetIndex == i)
                {
                    workbook.SetCurrentWorksheet(i);
                    workbook.CurrentWorksheet.SetHorizontalSplit(rowNumber, freeze, new Address(topLeftCellAddress), activePane);
                }
            }
            Worksheet givenWorksheet = WriteAndReadWorksheet(workbook, sheetIndex);
            assertRowSplit(rowNumber, freeze, givenWorksheet);
        }

        [Fact(DisplayName = "Test of the 'PaneSplitTopHeight' property defined by a split address with custom row heights, when writing and reading a worksheet")]
        public void PaneSplitTopHeightsWriteReadTest3()
        {
            Workbook workbook = PrepareWorkbook(4, "test");
            workbook.SetCurrentWorksheet(0);
            workbook.CurrentWorksheet.SetRowHeight(0, 18f);
            workbook.CurrentWorksheet.SetRowHeight(2, 22.5f);
            workbook.CurrentWorksheet.SetHorizontalSplit(4, false, new Address("D1"), Worksheet.WorksheetPane.topLeft);

            float expectedHeight = 0f;
            for (int i = 0; i < 4; i++)
            {
                if (workbook.CurrentWorksheet.RowHeights.ContainsKey(i))
                {
                    expectedHeight += Utils.GetInternalRowHeight(workbook.CurrentWorksheet.RowHeights[i]);
                }
                else
                {
                    expectedHeight += Utils.GetInternalRowHeight(Worksheet.DEFAULT_ROW_HEIGHT);
                }

            }
            Worksheet givenWorksheet = WriteAndReadWorksheet(workbook, 0);
            // There may be a deviation by rounding
            float delta = Math.Abs(expectedHeight - givenWorksheet.PaneSplitTopHeight.Value);
            Assert.True(delta < 0.15);
            Assert.Null(givenWorksheet.FreezeSplitPanes);
        }

        [Theory(DisplayName = "Test of the 'PaneSplitLeftWidth' property when writing and reading a worksheet")]
        [InlineData(27f, null, 0)]
        [InlineData(100f, null, 0)]
        [InlineData(10f, null, 0)]
        [InlineData(27f, Worksheet.WorksheetPane.topLeft, 0)]
        [InlineData(100f, Worksheet.WorksheetPane.topLeft, 0)]
        [InlineData(0f, Worksheet.WorksheetPane.topLeft, 0)]
        public void PaneSplitLeftWidthWriteReadTest(float width, Worksheet.WorksheetPane? activePane, int sheetIndex)
        {
            Workbook workbook = PrepareWorkbook(4, "test");
            for (int i = 0; i <= sheetIndex; i++)
            {
                if (sheetIndex == i)
                {
                    workbook.SetCurrentWorksheet(i);
                    workbook.CurrentWorksheet.SetVerticalSplit(width, new Address("A2"), activePane);
                }
            }
            Worksheet givenWorksheet = WriteAndReadWorksheet(workbook, sheetIndex);
            // There may be a deviation by rounding
            float delta = Math.Abs(width - givenWorksheet.PaneSplitLeftWidth.Value);
            Assert.True(delta < 0.1);
        }

        [Theory(DisplayName = "Test of the 'PaneSplitLeftWidth' property defined by a split address, when writing and reading a worksheet")]
        [InlineData(0, false, "A2", null, 0)]
        [InlineData(1, false, "B2", null, 0)]
        [InlineData(5, false, "G2", null, 0)]
        [InlineData(0, false, "A2", Worksheet.WorksheetPane.topLeft, 0)]
        [InlineData(1, false, "B2", Worksheet.WorksheetPane.bottomLeft, 0)]
        [InlineData(5, false, "G2", Worksheet.WorksheetPane.topRight, 0)]
        [InlineData(0, true, "A2", null, 0)]
        [InlineData(1, true, "B2", null, 0)]
        [InlineData(5, true, "G2", null, 0)]
        [InlineData(0, true, "A2", Worksheet.WorksheetPane.topLeft, 0)]
        [InlineData(1, true, "B2", Worksheet.WorksheetPane.bottomLeft, 0)]
        [InlineData(5, true, "G2", Worksheet.WorksheetPane.topRight, 0)]
        public void PaneSplitLeftWidthWriteReadTest2(int columnNumber, bool freeze, string topLeftCellAddress, Worksheet.WorksheetPane? activePane, int sheetIndex)
        {
            Workbook workbook = PrepareWorkbook(4, "test");
            for (int i = 0; i <= sheetIndex; i++)
            {
                if (sheetIndex == i)
                {
                    workbook.SetCurrentWorksheet(i);
                    workbook.CurrentWorksheet.SetVerticalSplit(columnNumber, freeze, new Address(topLeftCellAddress), activePane);
                }
            }
            Worksheet givenWorksheet = WriteAndReadWorksheet(workbook, sheetIndex);
            asserColumnSplit(columnNumber, freeze, givenWorksheet, false);
        }

        [Fact(DisplayName = "Test of the 'PaneSplitLeftWidth' property defined by a split address with custom column widths, when writing and reading a worksheet")]
        public void PaneSplitLeftWidthWriteReadTest3()
        {
            Workbook workbook = PrepareWorkbook(4, "test");
            workbook.SetCurrentWorksheet(0);
            workbook.CurrentWorksheet.SetColumnWidth(0, 18f);
            workbook.CurrentWorksheet.SetColumnWidth(2, 22.5f);
            workbook.CurrentWorksheet.SetVerticalSplit(4, false, new Address("D1"), Worksheet.WorksheetPane.topLeft);

            float expectedWidth = 0f;
            for(int i = 0; i < 4; i++)
            {
                if (workbook.CurrentWorksheet.Columns.ContainsKey(i))
                {
                    expectedWidth += Utils.GetInternalColumnWidth(workbook.CurrentWorksheet.Columns[i].Width);
                }
                else
                {
                    expectedWidth += Utils.GetInternalColumnWidth(Worksheet.DEFAULT_COLUMN_WIDTH);
                }
                
            }
            Worksheet givenWorksheet = WriteAndReadWorksheet(workbook, 0);
            // There may be a deviation by rounding
            float delta = Math.Abs(expectedWidth - givenWorksheet.PaneSplitLeftWidth.Value);
            Assert.True(delta < 0.15);
            Assert.Null(givenWorksheet.FreezeSplitPanes);
        }

        [Theory(DisplayName = "Test of the 'ActivePane' property when writing and reading a worksheet")]
        [InlineData(27f, null, 0)]
        [InlineData(100f, Worksheet.WorksheetPane.topLeft, 0)]
        [InlineData(0f, Worksheet.WorksheetPane.bottomLeft, 0)]
        [InlineData(27f, Worksheet.WorksheetPane.topRight, 0)]
        [InlineData(100f, Worksheet.WorksheetPane.bottomRight, 0)]
        public void PaneSplitActivePaneWriteReadTest(float height, Worksheet.WorksheetPane? activePane, int sheetIndex)
        {
            Workbook workbook = PrepareWorkbook(4, "test");
            for (int i = 0; i <= sheetIndex; i++)
            {
                if (sheetIndex == i)
                {
                    workbook.SetCurrentWorksheet(i);
                    workbook.CurrentWorksheet.SetHorizontalSplit(height, new Address("A2"), activePane);
                }
            }
            Worksheet givenWorksheet = WriteAndReadWorksheet(workbook, sheetIndex);
            Assert.Equal(activePane, givenWorksheet.ActivePane);
        }
        [Theory(DisplayName = "Test of the 'PaneSplitTopLeftCell' property when writing and reading a worksheet")]
        [InlineData(27f, null, "A1", 0)]
        [InlineData(100f, Worksheet.WorksheetPane.topLeft, "B2", 0)]
        [InlineData(0f, Worksheet.WorksheetPane.bottomLeft, "Z15", 0)]
        [InlineData(27f, Worksheet.WorksheetPane.topRight, "$A1", 0)]
        [InlineData(100f, Worksheet.WorksheetPane.bottomRight, "$D$4", 0)]
        public void PaneSplitTopLeftCellWriteReadTest(float height, Worksheet.WorksheetPane? activePane, string topLeftCellAddress, int sheetIndex)
        {
            Address topLeftCell = new Address(topLeftCellAddress);
            Workbook workbook = PrepareWorkbook(4, "test");
            for (int i = 0; i <= sheetIndex; i++)
            {
                if (sheetIndex == i)
                {
                    workbook.SetCurrentWorksheet(i);
                    workbook.CurrentWorksheet.SetHorizontalSplit(height, topLeftCell, activePane);
                }
            }
            Worksheet givenWorksheet = WriteAndReadWorksheet(workbook, sheetIndex);
            Assert.Equal(topLeftCell, givenWorksheet.PaneSplitTopLeftCell);
        }



        [Theory(DisplayName = "Test of the 'PaneSplitTopHeight' and 'PaneSplitLeftWidth' properties (combined X/Y-Split) when writing and reading a worksheet")]
        [InlineData(27f, 0f, null, 0)]
        [InlineData(100f, 0f, null, 0)]
        [InlineData(0f, 0f, null, 0)]
        [InlineData(27f, 27f, Worksheet.WorksheetPane.topLeft, 0)]
        [InlineData(100f, 27f, Worksheet.WorksheetPane.bottomLeft, 0)]
        [InlineData(0f, 27f, Worksheet.WorksheetPane.topRight, 0)]
        [InlineData(27f, 100f, null, 0)]
        [InlineData(100f, 100f, null, 0)]
        [InlineData(0f, 100f, null, 0)]
        [InlineData(27f, null, Worksheet.WorksheetPane.topLeft, 0)]
        [InlineData(100f, null, Worksheet.WorksheetPane.bottomLeft, 0)]
        [InlineData(0f, null, Worksheet.WorksheetPane.topRight, 0)]
        [InlineData(null, 100f, null, 0)]
        [InlineData(null, 27f, null, 0)]
        [InlineData(null, 0f, null, 0)]
        [InlineData(null, null, Worksheet.WorksheetPane.topLeft, 0)]
        public void PaneSplitWidthHeightWriteReadTest(float? width, float? height, Worksheet.WorksheetPane? activePane, int sheetIndex)
        {
            Workbook workbook = PrepareWorkbook(4, "test");
            for (int i = 0; i <= sheetIndex; i++)
            {
                if (sheetIndex == i)
                {
                    workbook.SetCurrentWorksheet(i);
                    workbook.CurrentWorksheet.SetSplit(width, height, new Address("B2"), activePane);
                }
            }
            Worksheet givenWorksheet = WriteAndReadWorksheet(workbook, sheetIndex);
            Assert.Equal(height, givenWorksheet.PaneSplitTopHeight);
            if (width == null)
            {
                Assert.Null(givenWorksheet.PaneSplitLeftWidth);
            }
            else
            {
                // There may be a deviation by rounding
                float delta = Math.Abs(width.Value - givenWorksheet.PaneSplitLeftWidth.Value);
                Assert.True(delta < 0.1);
            }
        }

        //public void PaneSplit

        [Theory(DisplayName = "Test of the'PaneSplitTopHeight' and the 'PaneSplitLeftWidth' properties (combined X/Y-Split) defined by a split address, when writing and reading a worksheet")]
        [InlineData(0, 0, false, "A2", null, 0)]
        [InlineData(1, 0, false, "B2", null, 0)]
        [InlineData(5, 0, false, "G2", null, 0)]
        [InlineData(0, 0, false, "A2", Worksheet.WorksheetPane.topLeft, 0)]
        [InlineData(1, 0, false, "B2", Worksheet.WorksheetPane.bottomLeft, 0)]
        [InlineData(5, 0, false, "G2", Worksheet.WorksheetPane.topRight, 0)]
        [InlineData(0, 1, true, "A2", null, 0)]
        [InlineData(1, 1, true, "B2", null, 0)]
        [InlineData(5, 1, true, "G2", null, 0)]
        [InlineData(0, 1, true, "A2", Worksheet.WorksheetPane.topLeft, 0)]
        [InlineData(1, 1, true, "B2", Worksheet.WorksheetPane.bottomLeft, 0)]
        [InlineData(5, 1, true, "G2", Worksheet.WorksheetPane.topRight, 0)]
        [InlineData(0, 15, true, "A20", null, 0)]
        [InlineData(1, 15, true, "B20", null, 0)]
        [InlineData(5, 15, true, "G20", null, 0)]
        [InlineData(0, 15, true, "A20", Worksheet.WorksheetPane.topLeft, 0)]
        [InlineData(1, 15, true, "B20", Worksheet.WorksheetPane.bottomLeft, 0)]
        [InlineData(5, 15, true, "G20", Worksheet.WorksheetPane.topRight, 0)]
        public void PaneSplitWidthHeightWriteReadTest2(int columnNumber, int rowNumber, bool freeze, string topLeftCellAddress, Worksheet.WorksheetPane? activePane, int sheetIndex)
        {
            Workbook workbook = PrepareWorkbook(4, "test");
            for (int i = 0; i <= sheetIndex; i++)
            {
                if (sheetIndex == i)
                {
                    workbook.SetCurrentWorksheet(i);
                    workbook.CurrentWorksheet.SetSplit(columnNumber, rowNumber, freeze, new Address(topLeftCellAddress), activePane);
                }
            }
            Worksheet givenWorksheet = WriteAndReadWorksheet(workbook, sheetIndex);
            asserColumnSplit(columnNumber, freeze, givenWorksheet, true);
            assertRowSplit(rowNumber, freeze, givenWorksheet);
        }

        private static void asserColumnSplit(int columnNumber, bool freeze, Worksheet givenWorksheet, bool xyApplied)
        {
            if (columnNumber == 0 && !xyApplied)
            {
                // No split at all (column 0)
                Assert.Null(givenWorksheet.PaneSplitAddress);
                Assert.Null(givenWorksheet.FreezeSplitPanes);
            }
            else
            {
                if (freeze)
                {
                    Assert.Equal(columnNumber, givenWorksheet.PaneSplitAddress.Value.Column);
                    Assert.Equal(freeze, givenWorksheet.FreezeSplitPanes.Value);
                }
                else
                {
                    float width = Utils.GetInternalColumnWidth(Worksheet.DEFAULT_COLUMN_WIDTH) * columnNumber;
                    if (width == 0)
                    {
                        // Not applied as x split
                        Assert.Null(givenWorksheet.PaneSplitLeftWidth);
                    }
                    else
                    {
                        // There may be a deviation by rounding
                        float delta = Math.Abs(width - givenWorksheet.PaneSplitLeftWidth.Value);
                        Assert.True(delta < 0.1);
                    }
                    Assert.Null(givenWorksheet.FreezeSplitPanes);
                }

            }
        }
        private static void assertRowSplit(int rowNumber, bool freeze, Worksheet givenWorksheet)
        {
            if (rowNumber == 0)
            {
                // No split at all (row 0)
                Assert.Null(givenWorksheet.PaneSplitAddress);
                Assert.Null(givenWorksheet.FreezeSplitPanes);
            }
            else
            {
                if (freeze)
                {
                    Assert.Equal(rowNumber, givenWorksheet.PaneSplitAddress.Value.Row);
                    Assert.Equal(freeze, givenWorksheet.FreezeSplitPanes.Value);
                }
                else
                {
                    float height = Worksheet.DEFAULT_ROW_HEIGHT * rowNumber;
                    Assert.Equal(height, givenWorksheet.PaneSplitTopHeight);
                    Assert.Null(givenWorksheet.FreezeSplitPanes);
                }
            }
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
