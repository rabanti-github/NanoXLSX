using NanoXLSX;
using NanoXLSX.Styles;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xunit;

namespace NanoXLSX_Test.Styles.WriteRead
{
    public class FillWriteReadTest
    {
        [Theory(DisplayName = "Test of the writing and reading of the foreground color style value")]
        [InlineData("FFAACC00", "test")]
        [InlineData("FFAADD00", 0.5f)]
        [InlineData("FFDDCC00", true)]
        [InlineData("FFAACCDD", null)]
        public void ForegroundColorTest(String color, object value)
        {
            Style style = new Style();
            style.CurrentFill.ForegroundColor = color;
            Cell cell = CreateWorkbook(value, style);

            Assert.Equal(color, cell.CellStyle.CurrentFill.ForegroundColor);
            Assert.NotEqual(Fill.PatternValue.none, cell.CellStyle.CurrentFill.PatternFill);
        }

        [Theory(DisplayName = "Test of the writing and reading of the background color style value")]
        [InlineData("FFAACC00", "test")]
        [InlineData("FFAADD00", 0.5f)]
        [InlineData("FFDDCC00", true)]
        [InlineData("FFAACCDD", null)]
        public void BackgroundColorTest(String color, object value)
        {
            Style style = new Style();
            style.CurrentFill.BackgroundColor = color;
            style.CurrentFill.PatternFill = Fill.PatternValue.darkGray;
            Cell cell = CreateWorkbook(value, style);

            Assert.Equal(color, cell.CellStyle.CurrentFill.BackgroundColor);
            Assert.Equal(Fill.PatternValue.darkGray, cell.CellStyle.CurrentFill.PatternFill);
        }

        private static Cell CreateWorkbook(object value, Style style)
        {
            Workbook workbook = new Workbook(false);
            workbook.AddWorksheet("sheet1");
            workbook.CurrentWorksheet.AddCell(value, "A1", style);
            MemoryStream stream = new MemoryStream();
            workbook.SaveAsStream(stream, true);
            stream.Position = 0;
            Workbook givenWorkbook = Workbook.Load(stream);
            Cell cell = givenWorkbook.CurrentWorksheet.Cells["A1"];
            Assert.Equal(value, cell.Value);
            return cell;
        }

    }
}
