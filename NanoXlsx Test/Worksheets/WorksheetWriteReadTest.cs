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
            Workbook workbook = new Workbook("worksheet1");
            workbook.AddWorksheet("worksheet2");
            workbook.AddWorksheet("worksheet3");
            workbook.AddWorksheet("worksheet4");
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
