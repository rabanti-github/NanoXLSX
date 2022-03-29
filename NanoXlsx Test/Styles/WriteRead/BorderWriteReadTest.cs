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

        public enum BorderDirection
        {
            Diagonal,
            Left,
            Right,
            Top,
            Bottom
        }

        [Theory(DisplayName = "Test of the writing and reading of the diagonal color style value")]
        [InlineData("FFAACC00", "test", true, true)]
        [InlineData("FFAADD00", 0.5f, true, false)]
        [InlineData("FFDDCC00", true, false, true)]
        [InlineData("FFAACCDD", null, false, false)]
        [InlineData(null, 22, true, true)]
        public void DiagonalColorTest(String color, object value, bool diagonalUp, bool diagonalDown)
        {
            Style style = new Style();
            style.CurrentBorder.DiagonalColor = color;
            style.CurrentBorder.DiagonalStyle = Border.StyleValue.dashDot;
            style.CurrentBorder.DiagonalUp = diagonalUp;
            style.CurrentBorder.DiagonalDown = diagonalDown;

            Cell cell = CreateWorkbook(value, style);

            Assert.Equal(color, cell.CellStyle.CurrentBorder.DiagonalColor);
            Assert.Equal(Border.StyleValue.dashDot, cell.CellStyle.CurrentBorder.DiagonalStyle);
            Assert.Equal(diagonalUp, cell.CellStyle.CurrentBorder.DiagonalUp);
            Assert.Equal(diagonalDown, cell.CellStyle.CurrentBorder.DiagonalDown);
        }

        [Theory(DisplayName = "Test of the writing and reading of the top color style value")]
        [InlineData("FFAACC00", "test")]
        [InlineData("FFAADD00", 0.5f)]
        [InlineData("FFDDCC00", true)]
        [InlineData("FFAACCDD", null)]
        [InlineData(null, 22)]
        public void TopColorTest(String color, object value)
        {
            Style style = new Style();
            style.CurrentBorder.TopColor = color;
            style.CurrentBorder.TopStyle = Border.StyleValue.s_double;

            Cell cell = CreateWorkbook(value, style);

            Assert.Equal(color, cell.CellStyle.CurrentBorder.TopColor);
            Assert.Equal(Border.StyleValue.s_double, cell.CellStyle.CurrentBorder.TopStyle);
        }


        [Theory(DisplayName = "Test of the writing and reading of the bottom color style value")]
        [InlineData("FFAACC00", "test")]
        [InlineData("FFAADD00", 0.5f)]
        [InlineData("FFDDCC00", true)]
        [InlineData("FFAACCDD", null)]
        [InlineData(null, 22)]
        public void BottomColorTest(String color, object value)
        {
            Style style = new Style();
            style.CurrentBorder.BottomColor = color;
            style.CurrentBorder.BottomStyle = Border.StyleValue.thin;

            Cell cell = CreateWorkbook(value, style);

            Assert.Equal(color, cell.CellStyle.CurrentBorder.BottomColor);
            Assert.Equal(Border.StyleValue.thin, cell.CellStyle.CurrentBorder.BottomStyle);
        }

        [Theory(DisplayName = "Test of the writing and reading of the left color style value")]
        [InlineData("FFAACC00", "test")]
        [InlineData("FFAADD00", 0.5f)]
        [InlineData("FFDDCC00", true)]
        [InlineData("FFAACCDD", null)]
        [InlineData(null, 22)]
        public void LeftColorTest(String color, object value)
        {
            Style style = new Style();
            style.CurrentBorder.LeftColor = color;
            style.CurrentBorder.LeftStyle = Border.StyleValue.dashDotDot;

            Cell cell = CreateWorkbook(value, style);

            Assert.Equal(color, cell.CellStyle.CurrentBorder.LeftColor);
            Assert.Equal(Border.StyleValue.dashDotDot, cell.CellStyle.CurrentBorder.LeftStyle);
        }

        [Theory(DisplayName = "Test of the writing and reading of the right color style value")]
        [InlineData("FFAACC00", "test")]
        [InlineData("FFAADD00", 0.5f)]
        [InlineData("FFDDCC00", true)]
        [InlineData("FFAACCDD", null)]
        [InlineData(null, 22)]
        public void RightColorTest(String color, object value)
        {
            Style style = new Style();
            style.CurrentBorder.RightColor = color;
            style.CurrentBorder.RightStyle = Border.StyleValue.dashed;

            Cell cell = CreateWorkbook(value, style);

            Assert.Equal(color, cell.CellStyle.CurrentBorder.RightColor);
            Assert.Equal(Border.StyleValue.dashed, cell.CellStyle.CurrentBorder.RightStyle);
        }

        [Theory(DisplayName = "Test of the writing and reading of border style value")]
        [InlineData(Border.StyleValue.dashDotDot, BorderDirection.Bottom)]
        [InlineData(Border.StyleValue.dashDot, BorderDirection.Top)]
        [InlineData(Border.StyleValue.dashed, BorderDirection.Left)]
        [InlineData(Border.StyleValue.dotted, BorderDirection.Right)]
        [InlineData(Border.StyleValue.hair, BorderDirection.Diagonal)]
        [InlineData(Border.StyleValue.medium, BorderDirection.Bottom)]
        [InlineData(Border.StyleValue.mediumDashDot, BorderDirection.Top)]
        [InlineData(Border.StyleValue.mediumDashDotDot, BorderDirection.Left)]
        [InlineData(Border.StyleValue.mediumDashed, BorderDirection.Right)]
        [InlineData(Border.StyleValue.slantDashDot, BorderDirection.Diagonal)]
        [InlineData(Border.StyleValue.thin, BorderDirection.Bottom)]
        [InlineData(Border.StyleValue.s_double, BorderDirection.Top)]
        [InlineData(Border.StyleValue.thick, BorderDirection.Left)]
        [InlineData(Border.StyleValue.none, BorderDirection.Right)]
        public void BorderStyleTest(Border.StyleValue styleValue, BorderDirection direction)
        {
            Style style = new Style();
            switch (direction)
            {
                case BorderDirection.Diagonal:
                    style.CurrentBorder.DiagonalStyle = styleValue;
                    break;
                case BorderDirection.Left:
                    style.CurrentBorder.LeftStyle = styleValue;
                    break;
                case BorderDirection.Right:
                    style.CurrentBorder.RightStyle = styleValue;
                    break;
                case BorderDirection.Top:
                    style.CurrentBorder.TopStyle = styleValue;
                    break;
                case BorderDirection.Bottom:
                    style.CurrentBorder.BottomStyle = styleValue;
                    break;
            }
            Cell cell = CreateWorkbook("test", style);
            switch (direction)
            {
                case BorderDirection.Diagonal:
                    Assert.Equal(styleValue, cell.CellStyle.CurrentBorder.DiagonalStyle);
                    break;
                case BorderDirection.Left:
                    Assert.Equal(styleValue, cell.CellStyle.CurrentBorder.LeftStyle);
                    break;
                case BorderDirection.Right:
                    Assert.Equal(styleValue, cell.CellStyle.CurrentBorder.RightStyle);
                    break;
                case BorderDirection.Top:
                    Assert.Equal(styleValue, cell.CellStyle.CurrentBorder.TopStyle);
                    break;
                case BorderDirection.Bottom:
                    Assert.Equal(styleValue, cell.CellStyle.CurrentBorder.BottomStyle);
                    break;
            }
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
