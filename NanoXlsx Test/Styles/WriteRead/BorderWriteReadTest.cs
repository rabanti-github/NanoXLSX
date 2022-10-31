using NanoXLSX;
using NanoXLSX.Shared.Enums.Styles;
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

        [Theory(DisplayName = "Test of the 'diagonal' value when writing and reading a Border style")]
        [InlineData("FFAACC00", "test", true, true)]
        [InlineData("FFAADD00", 0.5f, true, false)]
        [InlineData("FFDDCC00", true, false, true)]
        [InlineData("FFAACCDD", null, false, false)]
        [InlineData(null, 22, true, true)]
        public void DiagonalColorTest(string color, object value, bool diagonalUp, bool diagonalDown)
        {
            Style style = new Style();
            style.CurrentBorder.DiagonalColor = color;
            style.CurrentBorder.DiagonalStyle = BorderEnums.StyleValue.dashDot;
            style.CurrentBorder.DiagonalUp = diagonalUp;
            style.CurrentBorder.DiagonalDown = diagonalDown;

            Cell cell = TestUtils.SaveAndReadStyledCell(value, style, "A1");

            Assert.Equal(color, cell.CellStyle.CurrentBorder.DiagonalColor);
            Assert.Equal(BorderEnums.StyleValue.dashDot, cell.CellStyle.CurrentBorder.DiagonalStyle);
            Assert.Equal(diagonalUp, cell.CellStyle.CurrentBorder.DiagonalUp);
            Assert.Equal(diagonalDown, cell.CellStyle.CurrentBorder.DiagonalDown);
        }

        [Theory(DisplayName = "Test of the 'top' value when writing and reading a Border style")]
        [InlineData("FFAACC00", "test")]
        [InlineData("FFAADD00", 0.5f)]
        [InlineData("FFDDCC00", true)]
        [InlineData("FFAACCDD", null)]
        [InlineData(null, 22)]
        public void TopColorTest(string color, object value)
        {
            Style style = new Style();
            style.CurrentBorder.TopColor = color;
            style.CurrentBorder.TopStyle = BorderEnums.StyleValue.s_double;

            Cell cell = TestUtils.SaveAndReadStyledCell(value, style, "A1");

            Assert.Equal(color, cell.CellStyle.CurrentBorder.TopColor);
            Assert.Equal(BorderEnums.StyleValue.s_double, cell.CellStyle.CurrentBorder.TopStyle);
        }


        [Theory(DisplayName = "Test of the 'bottom' value when writing and reading a Border style")]
        [InlineData("FFAACC00", "test")]
        [InlineData("FFAADD00", 0.5f)]
        [InlineData("FFDDCC00", true)]
        [InlineData("FFAACCDD", null)]
        [InlineData(null, 22)]
        public void BottomColorTest(string color, object value)
        {
            Style style = new Style();
            style.CurrentBorder.BottomColor = color;
            style.CurrentBorder.BottomStyle = BorderEnums.StyleValue.thin;

            Cell cell = TestUtils.SaveAndReadStyledCell(value, style, "A1");

            Assert.Equal(color, cell.CellStyle.CurrentBorder.BottomColor);
            Assert.Equal(BorderEnums.StyleValue.thin, cell.CellStyle.CurrentBorder.BottomStyle);
        }

        [Theory(DisplayName = "Test of the 'left' value when writing and reading a Border style")]
        [InlineData("FFAACC00", "test")]
        [InlineData("FFAADD00", 0.5f)]
        [InlineData("FFDDCC00", true)]
        [InlineData("FFAACCDD", null)]
        [InlineData(null, 22)]
        public void LeftColorTest(string color, object value)
        {
            Style style = new Style();
            style.CurrentBorder.LeftColor = color;
            style.CurrentBorder.LeftStyle = BorderEnums.StyleValue.dashDotDot;

            Cell cell = TestUtils.SaveAndReadStyledCell(value, style, "A1");

            Assert.Equal(color, cell.CellStyle.CurrentBorder.LeftColor);
            Assert.Equal(BorderEnums.StyleValue.dashDotDot, cell.CellStyle.CurrentBorder.LeftStyle);
        }

        [Theory(DisplayName = "Test of the 'right' value when writing and reading a Border style")]
        [InlineData("FFAACC00", "test")]
        [InlineData("FFAADD00", 0.5f)]
        [InlineData("FFDDCC00", true)]
        [InlineData("FFAACCDD", null)]
        [InlineData(null, 22)]
        public void RightColorTest(string color, object value)
        {
            Style style = new Style();
            style.CurrentBorder.RightColor = color;
            style.CurrentBorder.RightStyle = BorderEnums.StyleValue.dashed;

            Cell cell = TestUtils.SaveAndReadStyledCell(value, style, "A1");

            Assert.Equal(color, cell.CellStyle.CurrentBorder.RightColor);
            Assert.Equal(BorderEnums.StyleValue.dashed, cell.CellStyle.CurrentBorder.RightStyle);
        }

        [Theory(DisplayName = "Test of the 'styleValue' property when writing and reading a Font style")]
        [InlineData(BorderEnums.StyleValue.dashDotDot, BorderDirection.Bottom)]
        [InlineData(BorderEnums.StyleValue.dashDot, BorderDirection.Top)]
        [InlineData(BorderEnums.StyleValue.dashed, BorderDirection.Left)]
        [InlineData(BorderEnums.StyleValue.dotted, BorderDirection.Right)]
        [InlineData(BorderEnums.StyleValue.hair, BorderDirection.Diagonal)]
        [InlineData(BorderEnums.StyleValue.medium, BorderDirection.Bottom)]
        [InlineData(BorderEnums.StyleValue.mediumDashDot, BorderDirection.Top)]
        [InlineData(BorderEnums.StyleValue.mediumDashDotDot, BorderDirection.Left)]
        [InlineData(BorderEnums.StyleValue.mediumDashed, BorderDirection.Right)]
        [InlineData(BorderEnums.StyleValue.slantDashDot, BorderDirection.Diagonal)]
        [InlineData(BorderEnums.StyleValue.thin, BorderDirection.Bottom)]
        [InlineData(BorderEnums.StyleValue.s_double, BorderDirection.Top)]
        [InlineData(BorderEnums.StyleValue.thick, BorderDirection.Left)]
        [InlineData(BorderEnums.StyleValue.none, BorderDirection.Right)]
        public void BorderStyleTest(BorderEnums.StyleValue styleValue, BorderDirection direction)
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
            Cell cell = TestUtils.SaveAndReadStyledCell("test", style, "A1");
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
    }
}
