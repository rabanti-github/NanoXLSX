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
        public void PaneSplitTopLeftCellWriteReadTest(float height, Worksheet.WorksheetPane? activePane, String topLeftCellAddress, int sheetIndex)
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
