using System;
using NanoXLSX.Exceptions;
using NanoXLSX.Styles;
using Xunit;
using static NanoXLSX.Styles.Border;

namespace NanoXLSX.Test.Core.StyleTest
{
    public class BorderTest
    {
        private Border exampleStyle;

        public BorderTest()
        {
            exampleStyle = new Border
            {
                BottomColor = "11001100",
                BottomStyle = StyleValue.DashDot,
                DiagonalColor = "8877AA00",
                DiagonalDown = true,
                DiagonalStyle = StyleValue.Thick,
                DiagonalUp = true,
                LeftColor = "9911DD00",
                LeftStyle = StyleValue.MediumDashDotDot,
                RightColor = "FF00AA00",
                RightStyle = StyleValue.DashDotDot,
                TopColor = "22222200",
                TopStyle = StyleValue.Dashed
            };
        }

        [Theory(DisplayName = "Test of the get and set function of the BottomColor property")]
        [InlineData("")]
        [InlineData(null)]
        [InlineData("FFAA3300")]
        public void BottomColorTest(string value)
        {
            Border border = new Border();
            Assert.Empty(border.BottomColor);
            border.BottomColor = value;
            Assert.Equal(value, border.BottomColor);
        }

        [Theory(DisplayName = "Test of the failing set function of the BottomColor property with invalid values")]
        [InlineData("77BB00")]
        [InlineData("0002200000")]
        [InlineData("XXXXXXXX")]
        public void BottomColorFailTest(string value)
        {
            Border border = new Border();
            Exception ex = Assert.Throws<StyleException>(() => border.BottomColor = value);
            Assert.Equal(typeof(StyleException), ex.GetType());
        }

        [Theory(DisplayName = "Test of the get and set function of the BottomStyle property")]
        [InlineData(StyleValue.DashDot)]
        [InlineData(StyleValue.DashDotDot)]
        [InlineData(StyleValue.Dashed)]
        [InlineData(StyleValue.Dotted)]
        [InlineData(StyleValue.Hair)]
        [InlineData(StyleValue.Medium)]
        [InlineData(StyleValue.MediumDashDot)]
        [InlineData(StyleValue.MediumDashDotDot)]
        [InlineData(StyleValue.MediumDashed)]
        [InlineData(StyleValue.None)]
        [InlineData(StyleValue.SlantDashDot)]
        [InlineData(StyleValue.Double)]
        [InlineData(StyleValue.Thick)]
        [InlineData(StyleValue.Thin)]
        public void BottomStyleTest(StyleValue value)
        {
            Border border = new Border();
            Assert.Equal(Border.DefaultBorderStyle, border.BottomStyle); // none is default
            border.BottomStyle = value;
            Assert.Equal(value, border.BottomStyle);
        }

        [Theory(DisplayName = "Test of the get and set function of the DiagonalColor property")]
        [InlineData("")]
        [InlineData(null)]
        [InlineData("FFAA3300")]
        public void DiagonalColorTest(string value)
        {
            Border border = new Border();
            Assert.Empty(border.DiagonalColor);
            border.DiagonalColor = value;
            Assert.Equal(value, border.DiagonalColor);
        }

        [Theory(DisplayName = "Test of the failing set function of the DiagonalColor property with invalid values")]
        [InlineData("77BB00")]
        [InlineData("0002200000")]
        [InlineData("XXXXXXXX")]
        public void DiagonalColorFailTest(string value)
        {
            Border border = new Border();
            Exception ex = Assert.Throws<StyleException>(() => border.DiagonalColor = value);
            Assert.Equal(typeof(StyleException), ex.GetType());
        }

        [Theory(DisplayName = "Test of the get and set function of the DiagonalStyle property")]
        [InlineData(StyleValue.DashDot)]
        [InlineData(StyleValue.DashDotDot)]
        [InlineData(StyleValue.Dashed)]
        [InlineData(StyleValue.Dotted)]
        [InlineData(StyleValue.Hair)]
        [InlineData(StyleValue.Medium)]
        [InlineData(StyleValue.MediumDashDot)]
        [InlineData(StyleValue.MediumDashDotDot)]
        [InlineData(StyleValue.MediumDashed)]
        [InlineData(StyleValue.None)]
        [InlineData(StyleValue.SlantDashDot)]
        [InlineData(StyleValue.Double)]
        [InlineData(StyleValue.Thick)]
        [InlineData(StyleValue.Thin)]
        public void DiagonalStyleTest(StyleValue value)
        {
            Border border = new Border();
            Assert.Equal(Border.DefaultBorderStyle, border.DiagonalStyle); // none is default
            border.DiagonalStyle = value;
            Assert.Equal(value, border.DiagonalStyle);
        }

        [Theory(DisplayName = "Test of the get and set function of the LeftColor property")]
        [InlineData("")]
        [InlineData(null)]
        [InlineData("FFAA3300")]
        public void LeftColorTest(string value)
        {
            Border border = new Border();
            Assert.Empty(border.LeftColor);
            border.LeftColor = value;
            Assert.Equal(value, border.LeftColor);
        }

        [Theory(DisplayName = "Test of the failing set function of the LeftColor property with invalid values")]
        [InlineData("77BB00")]
        [InlineData("0002200000")]
        [InlineData("XXXXXXXX")]
        public void LeftColorFailTest(string value)
        {
            Border border = new Border();
            Exception ex = Assert.Throws<StyleException>(() => border.LeftColor = value);
            Assert.Equal(typeof(StyleException), ex.GetType());
        }

        [Theory(DisplayName = "Test of the get and set function of the LeftColor property")]
        [InlineData(StyleValue.DashDot)]
        [InlineData(StyleValue.DashDotDot)]
        [InlineData(StyleValue.Dashed)]
        [InlineData(StyleValue.Dotted)]
        [InlineData(StyleValue.Hair)]
        [InlineData(StyleValue.Medium)]
        [InlineData(StyleValue.MediumDashDot)]
        [InlineData(StyleValue.MediumDashDotDot)]
        [InlineData(StyleValue.MediumDashed)]
        [InlineData(StyleValue.None)]
        [InlineData(StyleValue.SlantDashDot)]
        [InlineData(StyleValue.Double)]
        [InlineData(StyleValue.Thick)]
        [InlineData(StyleValue.Thin)]
        public void LeftStyleTest(StyleValue value)
        {
            Border border = new Border();
            Assert.Equal(Border.DefaultBorderStyle, border.LeftStyle); // none is default
            border.LeftStyle = value;
            Assert.Equal(value, border.LeftStyle);
        }

        [Theory(DisplayName = "Test of the get and set function of the RightColor property")]
        [InlineData("")]
        [InlineData(null)]
        [InlineData("FFAA3300")]
        public void RightColorTest(string value)
        {
            Border border = new Border();
            Assert.Empty(border.RightColor);
            border.RightColor = value;
            Assert.Equal(value, border.RightColor);
        }

        [Theory(DisplayName = "Test of the failing set function of the RightColor property with invalid values")]
        [InlineData("77BB00")]
        [InlineData("0002200000")]
        [InlineData("XXXXXXXX")]
        public void RightColorFailTest(string value)
        {
            Border border = new Border();
            Exception ex = Assert.Throws<StyleException>(() => border.RightColor = value);
            Assert.Equal(typeof(StyleException), ex.GetType());
        }

        [Theory(DisplayName = "Test of the get and set function of the RightStyle property")]
        [InlineData(StyleValue.DashDot)]
        [InlineData(StyleValue.DashDotDot)]
        [InlineData(StyleValue.Dashed)]
        [InlineData(StyleValue.Dotted)]
        [InlineData(StyleValue.Hair)]
        [InlineData(StyleValue.Medium)]
        [InlineData(StyleValue.MediumDashDot)]
        [InlineData(StyleValue.MediumDashDotDot)]
        [InlineData(StyleValue.MediumDashed)]
        [InlineData(StyleValue.None)]
        [InlineData(StyleValue.SlantDashDot)]
        [InlineData(StyleValue.Double)]
        [InlineData(StyleValue.Thick)]
        [InlineData(StyleValue.Thin)]
        public void RightStyleTest(StyleValue value)
        {
            Border border = new Border();
            Assert.Equal(Border.DefaultBorderStyle, border.RightStyle); // none is default
            border.RightStyle = value;
            Assert.Equal(value, border.RightStyle);
        }

        [Theory(DisplayName = "Test of the get and set function of the TopColor property")]
        [InlineData("")]
        [InlineData(null)]
        [InlineData("FFAA3300")]
        public void TopColorTest(string value)
        {
            Border border = new Border();
            Assert.Empty(border.TopColor);
            border.TopColor = value;
            Assert.Equal(value, border.TopColor);
        }

        [Theory(DisplayName = "Test of the failing set function of the TopColor property with invalid values")]
        [InlineData("77BB00")]
        [InlineData("0002200000")]
        [InlineData("XXXXXXXX")]
        public void TopColorFailTest(string value)
        {
            Border border = new Border();
            Exception ex = Assert.Throws<StyleException>(() => border.TopColor = value);
            Assert.Equal(typeof(StyleException), ex.GetType());
        }

        [Theory(DisplayName = "Test of the get and set function of the TopStyle property")]
        [InlineData(StyleValue.DashDot)]
        [InlineData(StyleValue.DashDotDot)]
        [InlineData(StyleValue.Dashed)]
        [InlineData(StyleValue.Dotted)]
        [InlineData(StyleValue.Hair)]
        [InlineData(StyleValue.Medium)]
        [InlineData(StyleValue.MediumDashDot)]
        [InlineData(StyleValue.MediumDashDotDot)]
        [InlineData(StyleValue.MediumDashed)]
        [InlineData(StyleValue.None)]
        [InlineData(StyleValue.SlantDashDot)]
        [InlineData(StyleValue.Double)]
        [InlineData(StyleValue.Thick)]
        [InlineData(StyleValue.Thin)]
        public void TopStyleTest(StyleValue value)
        {
            Border border = new Border();
            Assert.Equal(Border.DefaultBorderStyle, border.TopStyle); // none is default
            border.TopStyle = value;
            Assert.Equal(value, border.TopStyle);
        }

        [Fact(DisplayName = "Test of the CopyBorder function")]
        public void CopyBorderTest()
        {
            Border copy = exampleStyle.CopyBorder();
            Assert.Equal(exampleStyle.GetHashCode(), copy.GetHashCode());
        }

        [Fact(DisplayName = "Test of the Equals method")]
        public void EqualsTest()
        {
            Border style2 = (Border)exampleStyle.Copy();
            Assert.True(exampleStyle.Equals(style2));
        }

        [Fact(DisplayName = "Test of the Equals method (inequality of BottomColor)")]
        public void EqualsTest2()
        {
            Border style2 = (Border)exampleStyle.Copy();
            style2.BottomColor = string.Empty;
            Assert.False(exampleStyle.Equals(style2));
        }

        [Fact(DisplayName = "Test of the Equals method (inequality of BottomStyle)")]
        public void EqualsTest2b()
        {
            Border style2 = (Border)exampleStyle.Copy();
            style2.BottomStyle = StyleValue.Double;
            Assert.False(exampleStyle.Equals(style2));
        }

        [Fact(DisplayName = "Test of the Equals method (inequality of TopColor)")]
        public void EqualsTest2c()
        {
            Border style2 = (Border)exampleStyle.Copy();
            style2.TopColor = string.Empty;
            Assert.False(exampleStyle.Equals(style2));
        }

        [Fact(DisplayName = "Test of the Equals method (inequality of TopStyle)")]
        public void EqualsTest2d()
        {
            Border style2 = (Border)exampleStyle.Copy();
            style2.TopStyle = StyleValue.Double;
            Assert.False(exampleStyle.Equals(style2));
        }

        [Fact(DisplayName = "Test of the Equals method (inequality of LeftColor)")]
        public void EqualsTest2e()
        {
            Border style2 = (Border)exampleStyle.Copy();
            style2.LeftColor = string.Empty;
            Assert.False(exampleStyle.Equals(style2));
        }

        [Fact(DisplayName = "Test of the Equals method (inequality of LeftStyle)")]
        public void EqualsTest2f()
        {
            Border style2 = (Border)exampleStyle.Copy();
            style2.LeftStyle = StyleValue.Double;
            Assert.False(exampleStyle.Equals(style2));
        }

        [Fact(DisplayName = "Test of the Equals method (inequality of RightColor)")]
        public void EqualsTest2g()
        {
            Border style2 = (Border)exampleStyle.Copy();
            style2.RightColor = string.Empty;
            Assert.False(exampleStyle.Equals(style2));
        }

        [Fact(DisplayName = "Test of the Equals method (inequality of RightStyle)")]
        public void EqualsTest2h()
        {
            Border style2 = (Border)exampleStyle.Copy();
            style2.RightStyle = StyleValue.Double;
            Assert.False(exampleStyle.Equals(style2));
        }

        [Fact(DisplayName = "Test of the Equals method (inequality of DiagonalColor)")]
        public void EqualsTest2i()
        {
            Border style2 = (Border)exampleStyle.Copy();
            style2.DiagonalColor = string.Empty;
            Assert.False(exampleStyle.Equals(style2));
        }

        [Fact(DisplayName = "Test of the Equals method (inequality of DiagonalStyle)")]
        public void EqualsTest2j()
        {
            Border style2 = (Border)exampleStyle.Copy();
            style2.DiagonalStyle = StyleValue.Double;
            Assert.False(exampleStyle.Equals(style2));
        }

        [Fact(DisplayName = "Test of the Equals method (inequality of DiagonalDown)")]
        public void EqualsTest2k()
        {
            Border style2 = (Border)exampleStyle.Copy();
            style2.DiagonalDown = false;
            Assert.False(exampleStyle.Equals(style2));
        }

        [Fact(DisplayName = "Test of the Equals method (inequality of DiagonalUp)")]
        public void EqualsTest2l()
        {
            Border style2 = (Border)exampleStyle.Copy();
            style2.DiagonalUp = false;
            Assert.False(exampleStyle.Equals(style2));
        }

        [Theory(DisplayName = "Test of the Equals method (inequality on null or different objects)")]
        [InlineData(null)]
        [InlineData("text")]
        [InlineData(true)]
        public void EqualsTest3(object obj)
        {
            Assert.False(exampleStyle.Equals(obj));
        }

        [Theory(DisplayName = "Test of the Equals method when the origin object is null or not of the same type")]
        [InlineData(null)]
        [InlineData(true)]
        [InlineData("origin")]
        public void EqualsTest5(object origin)
        {
            Assert.False(exampleStyle.Equals(origin));
        }

        [Fact(DisplayName = "Test of the GetHashCode method (equality of two identical objects)")]
        public void GetHashCodeTest()
        {
            Border copy = (Border)exampleStyle.Copy();
            copy.InternalID = 99;  // Should not influence
            Assert.Equal(exampleStyle.GetHashCode(), copy.GetHashCode());
        }

        [Fact(DisplayName = "Test of the GetHashCode method (inequality of two different objects)")]
        public void GetHashCodeTest2()
        {
            Border copy = (Border)exampleStyle.Copy();
            copy.BottomColor = "AACCDD00";
            Assert.NotEqual(exampleStyle.GetHashCode(), copy.GetHashCode());
        }

        [Fact(DisplayName = "Test of the CompareTo method")]
        public void CompareToTest()
        {
            Border border = new Border();
            Border other = new Border();
            border.InternalID = null;
            other.InternalID = null;
            Assert.Equal(-1, border.CompareTo(other));
            border.InternalID = 5;
            Assert.Equal(1, border.CompareTo(other));
            Assert.Equal(1, border.CompareTo(null));
            other.InternalID = 5;
            Assert.Equal(0, border.CompareTo(other));
            other.InternalID = 4;
            Assert.Equal(1, border.CompareTo(other));
            other.InternalID = 6;
            Assert.Equal(-1, border.CompareTo(other));
        }

        // For code coverage
        [Fact(DisplayName = "Test of the ToString function")]
        public void ToStringTest()
        {
            Border border = new Border();
            string s1 = border.ToString();
            border.BottomColor = "FFAABBCC";
            Assert.NotEqual(s1, border.ToString()); // An explicit value comparison is probably not sensible
        }

    }
}
