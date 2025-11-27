using System;
using System.IO;
using NanoXLSX.Extensions;
using NanoXLSX.Utils;
using Xunit;

namespace NanoXLSX.Test.Writer_Reader.WorksheetTest
{
    public class ViewWriteReadTest
    {

        [Theory(DisplayName = "Test of the 'PaneSplitTopHeight' property when writing and reading a worksheet")]
        [InlineData(27f, null, 0)]
        [InlineData(100f, null, 0)]
        [InlineData(0f, null, 0)]
        [InlineData(27f, Worksheet.WorksheetPane.TopLeft, 0)]
        [InlineData(100f, Worksheet.WorksheetPane.BottomLeft, 0)]
        [InlineData(0f, Worksheet.WorksheetPane.TopRight, 0)]
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
        [InlineData(0, false, "A2", Worksheet.WorksheetPane.TopLeft, 0)]
        [InlineData(1, false, "A2", Worksheet.WorksheetPane.BottomLeft, 0)]
        [InlineData(15, false, "A18", Worksheet.WorksheetPane.TopRight, 0)]
        [InlineData(0, true, "A2", null, 0)]
        [InlineData(1, true, "A2", null, 0)]
        [InlineData(15, true, "A18", null, 0)]
        [InlineData(0, true, "A2", Worksheet.WorksheetPane.TopLeft, 0)]
        [InlineData(1, true, "A2", Worksheet.WorksheetPane.BottomLeft, 0)]
        [InlineData(15, true, "A18", Worksheet.WorksheetPane.TopRight, 0)]
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
            workbook.CurrentWorksheet.SetHorizontalSplit(4, false, new Address("D1"), Worksheet.WorksheetPane.TopLeft);

            float expectedHeight = 0f;
            for (int i = 0; i < 4; i++)
            {
                if (workbook.CurrentWorksheet.RowHeights.ContainsKey(i))
                {
                    expectedHeight += DataUtils.GetInternalRowHeight(workbook.CurrentWorksheet.RowHeights[i]);
                }
                else
                {
                    expectedHeight += DataUtils.GetInternalRowHeight(Worksheet.DefaultWorksheetRowHeight);
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
        [InlineData(27f, Worksheet.WorksheetPane.TopLeft, 0)]
        [InlineData(100f, Worksheet.WorksheetPane.TopLeft, 0)]
        [InlineData(0f, Worksheet.WorksheetPane.TopLeft, 0)]
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
        [InlineData(0, false, "A2", Worksheet.WorksheetPane.TopLeft, 0)]
        [InlineData(1, false, "B2", Worksheet.WorksheetPane.BottomLeft, 0)]
        [InlineData(5, false, "G2", Worksheet.WorksheetPane.TopRight, 0)]
        [InlineData(0, true, "A2", null, 0)]
        [InlineData(1, true, "B2", null, 0)]
        [InlineData(5, true, "G2", null, 0)]
        [InlineData(0, true, "A2", Worksheet.WorksheetPane.TopLeft, 0)]
        [InlineData(1, true, "B2", Worksheet.WorksheetPane.BottomLeft, 0)]
        [InlineData(5, true, "G2", Worksheet.WorksheetPane.TopRight, 0)]
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
            workbook.CurrentWorksheet.SetVerticalSplit(4, false, new Address("D1"), Worksheet.WorksheetPane.TopLeft);

            float expectedWidth = 0f;
            for (int i = 0; i < 4; i++)
            {
                if (workbook.CurrentWorksheet.Columns.ContainsKey(i))
                {
                    expectedWidth += DataUtils.GetInternalColumnWidth(workbook.CurrentWorksheet.Columns[i].Width);
                }
                else
                {
                    expectedWidth += DataUtils.GetInternalColumnWidth(Worksheet.DefaultWorksheetColumnWidth);
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
        [InlineData(100f, Worksheet.WorksheetPane.TopLeft, 0)]
        [InlineData(0f, Worksheet.WorksheetPane.BottomLeft, 0)]
        [InlineData(27f, Worksheet.WorksheetPane.TopRight, 0)]
        [InlineData(100f, Worksheet.WorksheetPane.BottomRight, 0)]
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
        [InlineData(100f, Worksheet.WorksheetPane.TopLeft, "B2", 0)]
        [InlineData(0f, Worksheet.WorksheetPane.BottomLeft, "Z15", 0)]
        [InlineData(27f, Worksheet.WorksheetPane.TopRight, "$A1", 0)]
        [InlineData(100f, Worksheet.WorksheetPane.BottomRight, "$D$4", 0)]
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
        [InlineData(27f, 27f, Worksheet.WorksheetPane.TopLeft, 0)]
        [InlineData(100f, 27f, Worksheet.WorksheetPane.BottomLeft, 0)]
        [InlineData(0f, 27f, Worksheet.WorksheetPane.TopRight, 0)]
        [InlineData(27f, 100f, null, 0)]
        [InlineData(100f, 100f, null, 0)]
        [InlineData(0f, 100f, null, 0)]
        [InlineData(27f, null, Worksheet.WorksheetPane.TopLeft, 0)]
        [InlineData(100f, null, Worksheet.WorksheetPane.BottomLeft, 0)]
        [InlineData(0f, null, Worksheet.WorksheetPane.TopRight, 0)]
        [InlineData(null, 100f, null, 0)]
        [InlineData(null, 27f, null, 0)]
        [InlineData(null, 0f, null, 0)]
        [InlineData(null, null, Worksheet.WorksheetPane.TopLeft, 0)]
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
        [InlineData(0, 0, false, "A2", Worksheet.WorksheetPane.TopLeft, 0)]
        [InlineData(1, 0, false, "B2", Worksheet.WorksheetPane.BottomLeft, 0)]
        [InlineData(5, 0, false, "G2", Worksheet.WorksheetPane.TopRight, 0)]
        [InlineData(0, 1, true, "A2", null, 0)]
        [InlineData(1, 1, true, "B2", null, 0)]
        [InlineData(5, 1, true, "G2", null, 0)]
        [InlineData(0, 1, true, "A2", Worksheet.WorksheetPane.TopLeft, 0)]
        [InlineData(1, 1, true, "B2", Worksheet.WorksheetPane.BottomLeft, 0)]
        [InlineData(5, 1, true, "G2", Worksheet.WorksheetPane.TopRight, 0)]
        [InlineData(0, 15, true, "A20", null, 0)]
        [InlineData(1, 15, true, "B20", null, 0)]
        [InlineData(5, 15, true, "G20", null, 0)]
        [InlineData(0, 15, true, "A20", Worksheet.WorksheetPane.TopLeft, 0)]
        [InlineData(1, 15, true, "B20", Worksheet.WorksheetPane.BottomLeft, 0)]
        [InlineData(5, 15, true, "G20", Worksheet.WorksheetPane.TopRight, 0)]
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

        [Theory(DisplayName = "Test of the 'ShowGridLines' property, when writing and reading a worksheet")]
        [InlineData(true, 0)]
        [InlineData(false, 0)]
        [InlineData(true, 2)]
        [InlineData(false, 2)]
        public void ShowGridLinesWriteReadTest(bool showGridLines, int sheetIndex)
        {
            Workbook workbook = PrepareWorkbook(4, "test");
            workbook.SetCurrentWorksheet(sheetIndex);
            workbook.CurrentWorksheet.ShowGridLines = showGridLines;
            Worksheet givenWorksheet = WriteAndReadWorksheet(workbook, sheetIndex);
            Assert.Equal(showGridLines, givenWorksheet.ShowGridLines);
        }

        [Theory(DisplayName = "Test of the 'ShowRowColumnHeaders' property, when writing and reading a worksheet")]
        [InlineData(true, 0)]
        [InlineData(false, 0)]
        [InlineData(true, 2)]
        [InlineData(false, 2)]
        public void ShowRowColumnHeadersWriteReadTest(bool showRowColumnHeaders, int sheetIndex)
        {
            Workbook workbook = PrepareWorkbook(4, "test");
            workbook.SetCurrentWorksheet(sheetIndex);
            workbook.CurrentWorksheet.ShowRowColumnHeaders = showRowColumnHeaders;
            Worksheet givenWorksheet = WriteAndReadWorksheet(workbook, sheetIndex);
            Assert.Equal(showRowColumnHeaders, givenWorksheet.ShowRowColumnHeaders);
        }

        [Theory(DisplayName = "Test of the 'ShowRuler' property, when writing and reading a worksheet")]
        [InlineData(true, true, Worksheet.SheetViewType.PageLayout, 0)]
        [InlineData(false, true, Worksheet.SheetViewType.PageBreakPreview, 0)]
        [InlineData(true, true, Worksheet.SheetViewType.Normal, 2)]
        [InlineData(false, false, Worksheet.SheetViewType.PageLayout, 2)]
        [InlineData(true, true, Worksheet.SheetViewType.PageBreakPreview, 2)]
        [InlineData(false, true, Worksheet.SheetViewType.Normal, 1)]
        public void ShowRulerWriteReadTest(bool showRuler, bool expectedShowRuler, Worksheet.SheetViewType viewType, int sheetIndex)
        {
            Workbook workbook = PrepareWorkbook(4, "test");
            workbook.SetCurrentWorksheet(sheetIndex);
            workbook.CurrentWorksheet.ViewType = viewType;
            workbook.CurrentWorksheet.ShowRuler = showRuler;
            Worksheet givenWorksheet = WriteAndReadWorksheet(workbook, sheetIndex);
            Assert.Equal(viewType, givenWorksheet.ViewType);
            Assert.Equal(expectedShowRuler, givenWorksheet.ShowRuler);
        }

        [Theory(DisplayName = "Test of the 'ViewType' property, when writing and reading a worksheet")]
        [InlineData(Worksheet.SheetViewType.PageLayout, 0)]
        [InlineData(Worksheet.SheetViewType.PageBreakPreview, 0)]
        [InlineData(Worksheet.SheetViewType.Normal, 0)]
        [InlineData(Worksheet.SheetViewType.PageLayout, 2)]
        [InlineData(Worksheet.SheetViewType.PageBreakPreview, 2)]
        [InlineData(Worksheet.SheetViewType.Normal, 2)]
        public void ViewTypeWriteReadTest(Worksheet.SheetViewType viewType, int sheetIndex)
        {
            Workbook workbook = PrepareWorkbook(4, "test");
            workbook.SetCurrentWorksheet(sheetIndex);
            workbook.CurrentWorksheet.ViewType = viewType;
            Worksheet givenWorksheet = WriteAndReadWorksheet(workbook, sheetIndex);
            Assert.Equal(viewType, givenWorksheet.ViewType);
        }

        [Theory(DisplayName = "Test of the 'ZoomFactor' property, when writing and reading a worksheet")]
        [InlineData(Worksheet.SheetViewType.Normal, 0, 0)]
        [InlineData(Worksheet.SheetViewType.Normal, 10, 2)]
        [InlineData(Worksheet.SheetViewType.Normal, 100, 0)]
        [InlineData(Worksheet.SheetViewType.Normal, 250, 2)]
        [InlineData(Worksheet.SheetViewType.Normal, 400, 0)]
        [InlineData(Worksheet.SheetViewType.PageLayout, 0, 2)]
        [InlineData(Worksheet.SheetViewType.PageLayout, 10, 0)]
        [InlineData(Worksheet.SheetViewType.PageLayout, 100, 2)]
        [InlineData(Worksheet.SheetViewType.PageLayout, 250, 0)]
        [InlineData(Worksheet.SheetViewType.PageLayout, 400, 2)]
        [InlineData(Worksheet.SheetViewType.PageBreakPreview, 0, 0)]
        [InlineData(Worksheet.SheetViewType.PageBreakPreview, 10, 2)]
        [InlineData(Worksheet.SheetViewType.PageBreakPreview, 100, 0)]
        [InlineData(Worksheet.SheetViewType.PageBreakPreview, 250, 2)]
        [InlineData(Worksheet.SheetViewType.PageBreakPreview, 400, 0)]
        public void ZoomFactorWriteReadTest(Worksheet.SheetViewType viewType, int zoomFactor, int sheetIndex)
        {
            Workbook workbook = PrepareWorkbook(4, "test");
            workbook.SetCurrentWorksheet(sheetIndex);
            workbook.CurrentWorksheet.ViewType = viewType;
            workbook.CurrentWorksheet.ZoomFactor = zoomFactor;
            Worksheet givenWorksheet = WriteAndReadWorksheet(workbook, sheetIndex);
            Assert.Equal(viewType, givenWorksheet.ViewType);
            Assert.Equal(zoomFactor, givenWorksheet.ZoomFactor);
        }

        [Theory(DisplayName = "Test of the 'SetZoomFactor' function, when writing and reading a worksheet")]
        [InlineData(Worksheet.SheetViewType.PageLayout, Worksheet.SheetViewType.Normal, 0, 0)]
        [InlineData(Worksheet.SheetViewType.PageBreakPreview, Worksheet.SheetViewType.Normal, 10, 2)]
        [InlineData(Worksheet.SheetViewType.PageLayout, Worksheet.SheetViewType.Normal, 100, 0)]
        [InlineData(Worksheet.SheetViewType.PageBreakPreview, Worksheet.SheetViewType.Normal, 250, 2)]
        [InlineData(Worksheet.SheetViewType.PageLayout, Worksheet.SheetViewType.Normal, 400, 0)]
        [InlineData(Worksheet.SheetViewType.Normal, Worksheet.SheetViewType.PageLayout, 0, 2)]
        [InlineData(Worksheet.SheetViewType.PageBreakPreview, Worksheet.SheetViewType.PageLayout, 10, 0)]
        [InlineData(Worksheet.SheetViewType.Normal, Worksheet.SheetViewType.PageLayout, 100, 2)]
        [InlineData(Worksheet.SheetViewType.PageBreakPreview, Worksheet.SheetViewType.PageLayout, 250, 0)]
        [InlineData(Worksheet.SheetViewType.Normal, Worksheet.SheetViewType.PageLayout, 400, 2)]
        [InlineData(Worksheet.SheetViewType.Normal, Worksheet.SheetViewType.PageBreakPreview, 0, 0)]
        [InlineData(Worksheet.SheetViewType.PageLayout, Worksheet.SheetViewType.PageBreakPreview, 10, 2)]
        [InlineData(Worksheet.SheetViewType.Normal, Worksheet.SheetViewType.PageBreakPreview, 100, 0)]
        [InlineData(Worksheet.SheetViewType.PageLayout, Worksheet.SheetViewType.PageBreakPreview, 250, 2)]
        [InlineData(Worksheet.SheetViewType.Normal, Worksheet.SheetViewType.PageBreakPreview, 400, 0)]
        public void SetZoomFactorWriteReadTest(Worksheet.SheetViewType initialViewType, Worksheet.SheetViewType additionalViewType, int zoomFactor, int sheetIndex)
        {
            Workbook workbook = PrepareWorkbook(4, "test");
            workbook.SetCurrentWorksheet(sheetIndex);
            workbook.CurrentWorksheet.ViewType = initialViewType;
            workbook.CurrentWorksheet.SetZoomFactor(additionalViewType, zoomFactor);
            workbook.SaveAs("c:\\purge-temp\\testZoom.xlsx");
            Worksheet givenWorksheet = WriteAndReadWorksheet(workbook, sheetIndex);
            if (initialViewType != Worksheet.SheetViewType.Normal && additionalViewType != Worksheet.SheetViewType.Normal)
            {
                Assert.Equal(3, givenWorksheet.ZoomFactors.Count);
                Assert.Equal(100, givenWorksheet.ZoomFactors[Worksheet.SheetViewType.Normal]);
            }
            else
            {
                Assert.Equal(2, givenWorksheet.ZoomFactors.Count);
            }
            Assert.Equal(zoomFactor, givenWorksheet.ZoomFactors[additionalViewType]);
            Assert.Equal(100, givenWorksheet.ZoomFactors[initialViewType]);
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
                    float width = DataUtils.GetInternalColumnWidth(Worksheet.DefaultWorksheetColumnWidth) * columnNumber;
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
                    float height = Worksheet.DefaultWorksheetRowHeight * rowNumber;
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
                Workbook readWorkbook = WorkbookReader.Load(stream);
                return readWorkbook.Worksheets[worksheetIndex];
            }
        }
    }
}
