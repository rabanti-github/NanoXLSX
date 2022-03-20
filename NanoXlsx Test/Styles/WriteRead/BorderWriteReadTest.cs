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
    public class BorderWriteReadTest
    {

        private enum BorderDirection
        {
            Diagonal,
            Left,
            Right,
            Top,
            Bottom
        }

        /*
        private string diagonalColor;
        private string leftColor;
        private string rightColor;
        private string topColor;
        private string bottomColor;
         */

        [Theory(DisplayName = "Test of the writing and reading of the diagonal color style value")]
        [InlineData("FFAACC00", "test")]
        [InlineData("FFAADD00", 0.5f)]
        [InlineData("FFDDCC00", true)]
        [InlineData("FFAACCDD", null)]
        public void DiagonalColorTest(String color, object value)
        {
            Style style = new Style();
            style.CurrentBorder.DiagonalColor = color;
            style.CurrentBorder.DiagonalStyle = Border.StyleValue.dashDot;
            style.CurrentBorder.DiagonalUp = true;

            Cell cell = CreateWorkbook(value, style);

            Assert.Equal(color, cell.CellStyle.CurrentBorder.DiagonalColor);
            Assert.Equal(Border.StyleValue.dashDot, cell.CellStyle.CurrentBorder.DiagonalStyle);
            Assert.True(cell.CellStyle.CurrentBorder.DiagonalUp);
            Assert.False(cell.CellStyle.CurrentBorder.DiagonalDown);
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
