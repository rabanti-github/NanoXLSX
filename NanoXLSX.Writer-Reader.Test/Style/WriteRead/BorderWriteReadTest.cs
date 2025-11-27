using NanoXLSX.Styles;
using NanoXLSX.Test.Writer_Reader.Utils;
using Xunit;
using static NanoXLSX.Styles.Border;

namespace NanoXLSX.Test.Writer_Reader.StyleTest
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
        [InlineData("", 22, true, true)]
        public void DiagonalColorTest(string color, object value, bool diagonalUp, bool diagonalDown)
        {
            Style style = new Style();
            style.CurrentBorder.DiagonalColor = color;
            style.CurrentBorder.DiagonalStyle = StyleValue.DashDot;
            style.CurrentBorder.DiagonalUp = diagonalUp;
            style.CurrentBorder.DiagonalDown = diagonalDown;

            Cell cell = TestUtils.SaveAndReadStyledCell(value, style, "A1");

            Assert.Equal(color, cell.CellStyle.CurrentBorder.DiagonalColor);
            Assert.Equal(StyleValue.DashDot, cell.CellStyle.CurrentBorder.DiagonalStyle);
            Assert.Equal(diagonalUp, cell.CellStyle.CurrentBorder.DiagonalUp);
            Assert.Equal(diagonalDown, cell.CellStyle.CurrentBorder.DiagonalDown);
        }

        [Theory(DisplayName = "Test of the 'top' value when writing and reading a Border style")]
        [InlineData("FFAACC00", "test")]
        [InlineData("FFAADD00", 0.5f)]
        [InlineData("FFDDCC00", true)]
        [InlineData("FFAACCDD", null)]
        [InlineData("", 22)]
        public void TopColorTest(string color, object value)
        {
            Style style = new Style();
            style.CurrentBorder.TopColor = color;
            style.CurrentBorder.TopStyle = StyleValue.Double;

            Cell cell = TestUtils.SaveAndReadStyledCell(value, style, "A1");

            Assert.Equal(color, cell.CellStyle.CurrentBorder.TopColor);
            Assert.Equal(StyleValue.Double, cell.CellStyle.CurrentBorder.TopStyle);
        }


        [Theory(DisplayName = "Test of the 'bottom' value when writing and reading a Border style")]
        [InlineData("FFAACC00", "test")]
        [InlineData("FFAADD00", 0.5f)]
        [InlineData("FFDDCC00", true)]
        [InlineData("FFAACCDD", null)]
        [InlineData("", 22)]
        public void BottomColorTest(string color, object value)
        {
            Style style = new Style();
            style.CurrentBorder.BottomColor = color;
            style.CurrentBorder.BottomStyle = StyleValue.Thin;

            Cell cell = TestUtils.SaveAndReadStyledCell(value, style, "A1");

            Assert.Equal(color, cell.CellStyle.CurrentBorder.BottomColor);
            Assert.Equal(StyleValue.Thin, cell.CellStyle.CurrentBorder.BottomStyle);
        }

        [Theory(DisplayName = "Test of the 'left' value when writing and reading a Border style")]
        [InlineData("FFAACC00", "test")]
        [InlineData("FFAADD00", 0.5f)]
        [InlineData("FFDDCC00", true)]
        [InlineData("FFAACCDD", null)]
        [InlineData("", 22)]
        public void LeftColorTest(string color, object value)
        {
            Style style = new Style();
            style.CurrentBorder.LeftColor = color;
            style.CurrentBorder.LeftStyle = StyleValue.DashDotDot;

            Cell cell = TestUtils.SaveAndReadStyledCell(value, style, "A1");

            Assert.Equal(color, cell.CellStyle.CurrentBorder.LeftColor);
            Assert.Equal(StyleValue.DashDotDot, cell.CellStyle.CurrentBorder.LeftStyle);
        }

        [Theory(DisplayName = "Test of the 'right' value when writing and reading a Border style")]
        [InlineData("FFAACC00", "test")]
        [InlineData("FFAADD00", 0.5f)]
        [InlineData("FFDDCC00", true)]
        [InlineData("FFAACCDD", null)]
        [InlineData("", 22)]
        public void RightColorTest(string color, object value)
        {
            Style style = new Style();
            style.CurrentBorder.RightColor = color;
            style.CurrentBorder.RightStyle = StyleValue.Dashed;

            Cell cell = TestUtils.SaveAndReadStyledCell(value, style, "A1");

            Assert.Equal(color, cell.CellStyle.CurrentBorder.RightColor);
            Assert.Equal(StyleValue.Dashed, cell.CellStyle.CurrentBorder.RightStyle);
        }

        [Theory(DisplayName = "Test of the 'styleValue' property when writing and reading a Font style")]
        [InlineData(StyleValue.DashDotDot, BorderDirection.Bottom)]
        [InlineData(StyleValue.DashDot, BorderDirection.Top)]
        [InlineData(StyleValue.Dashed, BorderDirection.Left)]
        [InlineData(StyleValue.Dotted, BorderDirection.Right)]
        [InlineData(StyleValue.Hair, BorderDirection.Diagonal)]
        [InlineData(StyleValue.Medium, BorderDirection.Bottom)]
        [InlineData(StyleValue.MediumDashDot, BorderDirection.Top)]
        [InlineData(StyleValue.MediumDashDotDot, BorderDirection.Left)]
        [InlineData(StyleValue.MediumDashed, BorderDirection.Right)]
        [InlineData(StyleValue.SlantDashDot, BorderDirection.Diagonal)]
        [InlineData(StyleValue.Thin, BorderDirection.Bottom)]
        [InlineData(StyleValue.Double, BorderDirection.Top)]
        [InlineData(StyleValue.Thick, BorderDirection.Left)]
        [InlineData(StyleValue.None, BorderDirection.Right)]
        public void BorderStyleTest(StyleValue styleValue, BorderDirection direction)
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
