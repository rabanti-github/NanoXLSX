using NanoXLSX.Exceptions;
using NanoXLSX.Styles;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xunit;

namespace NanoXLSX_Test.Styles
{
    public class FontTest
    {

        private Font exampleStyle;

        public FontTest()
        {
            exampleStyle = new Font();
            exampleStyle.Bold = true;
            exampleStyle.Italic = true;
            exampleStyle.Underline = true;
            exampleStyle.DoubleUnderline = true;
            exampleStyle.Strike = true;
            exampleStyle.Charset = "ASCII";
            exampleStyle.Size = 15;
            exampleStyle.Name = "Arial";
            exampleStyle.Family = "X";
            exampleStyle.ColorTheme = 10;
            exampleStyle.ColorValue = "FF22AACC";
            exampleStyle.Scheme = Font.SchemeValue.minor;
            exampleStyle.VerticalAlign = Font.VerticalAlignValue.subscript;
        }


        [Fact(DisplayName = "Test of the default values")]
        public void DefaultValuesTest()
        {
            Assert.Equal(11f, Font.DEFAULT_FONT_SIZE);
            Assert.Equal("2", Font.DEFAULT_FONT_FAMILY);
            Assert.Equal(Font.SchemeValue.minor, Font.DEFAULT_FONT_SCHEME);
            Assert.Equal(Font.VerticalAlignValue.none, Font.DEFAULT_VERTICAL_ALIGN);
            Assert.Equal("Calibri", Font.DEFAULT_FONT_NAME);
        }


        [Fact(DisplayName = "Test of the constructor")]
        public void ConstructorTest()
        {
            Font font = new Font();
            Assert.Equal(Font.DEFAULT_FONT_SIZE, font.Size);
            Assert.Equal(Font.DEFAULT_FONT_NAME, font.Name);
            Assert.Equal(Font.DEFAULT_FONT_FAMILY, font.Family);
            Assert.Equal(Font.DEFAULT_FONT_SCHEME, font.Scheme);
            Assert.Equal(Font.DEFAULT_VERTICAL_ALIGN, font.VerticalAlign);
            Assert.Equal("", font.ColorValue);
            Assert.Equal("", font.Charset);
            Assert.Equal(1, font.ColorTheme);
        }


        [Theory(DisplayName = "Test of the get and set function of the Bold property")]
        [InlineData(true)]
        [InlineData(false)]
        public void BoldTest(bool value)
        {
            Font font = new Font();
            Assert.False(font.Bold);
            font.Bold = value;
            Assert.Equal(value, font.Bold);
        }

        [Theory(DisplayName = "Test of the get and set function of the Italic property")]
        [InlineData(true)]
        [InlineData(false)]
        public void ItalicTest(bool value)
        {
            Font font = new Font();
            Assert.False(font.Italic);
            font.Italic = value;
            Assert.Equal(value, font.Italic);
        }

        [Theory(DisplayName = "Test of the get and set function of the Underline property")]
        [InlineData(true)]
        [InlineData(false)]
        public void UnderlineTest(bool value)
        {
            Font font = new Font();
            Assert.False(font.Underline);
            font.Underline = value;
            Assert.Equal(value, font.Underline);
        }

        [Theory(DisplayName = "Test of the get and set function of the DoubleUnderline property")]
        [InlineData(true)]
        [InlineData(false)]
        public void DoubleUnderlineTest(bool value)
        {
            Font font = new Font();
            Assert.False(font.DoubleUnderline);
            font.DoubleUnderline = value;
            Assert.Equal(value, font.DoubleUnderline);
        }

        [Theory(DisplayName = "Test of the get and set function of the Strike property")]
        [InlineData(true)]
        [InlineData(false)]
        public void StrikeTest(bool value)
        {
            Font font = new Font();
            Assert.False(font.Strike);
            font.Strike = value;
            Assert.Equal(value, font.Strike);
        }

        [Theory(DisplayName = "Test of the get and set function of the Charset property")]
        [InlineData("")]
        [InlineData("ASCII")]
        public void CharsetTest(string value)
        {
            Font font = new Font();
            Assert.Equal(string.Empty, font.Charset);
            font.Charset = value;
            Assert.Equal(value, font.Charset);
        }

        [Theory(DisplayName = "Test of the get and set function of the Size property")]
        [InlineData(8)]
        [InlineData(75)]
        [InlineData(11)]
        public void SizeTest(int value)
        {
            Font font = new Font();
            Assert.Equal(Font.DEFAULT_FONT_SIZE, font.Size); // 11 is default
            font.Size = value;
            Assert.Equal(value, font.Size);
        }

        [Theory(DisplayName = "Test of the auto-adjusting set function of the Size property (invalid values)")]
        [InlineData(0f, 1f)]
        [InlineData(7f, 7f)]
        [InlineData(-100f, 1f)]
        [InlineData(0.5f, 1f)]
        [InlineData(200f, 200f)]
        [InlineData(500f, 409f)]
        [InlineData(409.05f, 409f)]
        public void SizeFailTest(float givenValue, float expectedValue)
        {
            Font font = new Font();
            font.Size = givenValue;
            Assert.Equal(expectedValue, font.Size);
        }

        [Theory(DisplayName = "Test of the get and set function of the Name property")]
        [InlineData("Calibri")]
        [InlineData("Arial")]
        [InlineData("---")] // Not a font but a valid string
        public void NameTest(string value)
        {
            Font font = new Font();
            Assert.Equal(Font.DEFAULT_FONT_NAME, font.Name); // Default is 'Calibri'
            font.Name = value;
            Assert.Equal(value, font.Name);
        }

        [Fact(DisplayName = "Test of the failing set function of the Name property")]
        public void NameFailTest()
        {
            Font font = new Font();
            Assert.Throws<StyleException>(() => font.Name = null);
            Assert.Throws<StyleException>(() => font.Name = "");
        }

        [Theory(DisplayName = "Test of the get and set function of the Family property")]
        [InlineData("")]
        [InlineData("4")]
        public void FamilyTest(string value)
        {
            Font font = new Font();
            Assert.Equal(Font.DEFAULT_FONT_FAMILY, font.Family);
            font.Family = value;
            Assert.Equal(value, font.Family);
        }

        [Theory(DisplayName = "Test of the get and set function of the ColorTheme property")]
        [InlineData(1)]
        [InlineData(10)]
        public void ColorThemeTest(int value)
        {
            Font font = new Font();
            Assert.Equal(1, font.ColorTheme); // 1 is default
            font.ColorTheme = value;
            Assert.Equal(value, font.ColorTheme);
        }

        [Theory(DisplayName = "Test of the failing set function of the ColorTheme property (invalid values)")]
        [InlineData(0)]
        [InlineData(-100)]
        public void ColorThemeFailTest(int value)
        {
            Font font = new Font();
            Exception ex = Assert.Throws<StyleException>(() => font.ColorTheme = value);
            Assert.Equal(typeof(StyleException), ex.GetType());
        }

        [Theory(DisplayName = "Test of the get and set function of the ColorValue property")]
        [InlineData("")]
        [InlineData(null)]
        [InlineData("FFAA22CC")]
        public void ColorValueTest(string value)
        {
            Font font = new Font();
            Assert.Equal(string.Empty, font.ColorValue); // default is empty
            font.ColorValue = value;
            Assert.Equal(value, font.ColorValue);
        }

        [Theory(DisplayName = "Test of the failing set function of the ColorValue property (invalid values)")]
        [InlineData("77BB00")]
        [InlineData("0002200000")]
        [InlineData("XXXXXXXX")]
        public void ColorValueFailTest(string value)
        {
            Font font = new Font();
            Exception ex = Assert.Throws<StyleException>(() => font.ColorValue = value);
            Assert.Equal(typeof(StyleException), ex.GetType());
        }

        [Theory(DisplayName = "Test of the get and set function of the Scheme property")]
        [InlineData(Font.SchemeValue.major)]
        [InlineData(Font.SchemeValue.minor)]
        [InlineData(Font.SchemeValue.none)]
        public void SchmeTest(Font.SchemeValue value)
        {
            Font font = new Font();
            Assert.Equal(Font.DEFAULT_FONT_SCHEME, font.Scheme); // default is minor
            font.Scheme = value;
            Assert.Equal(value, font.Scheme);
        }

        [Theory(DisplayName = "Test of the get and set function of the VerticalAlign property")]
        [InlineData(Font.VerticalAlignValue.none)]
        [InlineData(Font.VerticalAlignValue.subscript)]
        [InlineData(Font.VerticalAlignValue.superscript)]
        public void VerticalAlignTest(Font.VerticalAlignValue value)
        {
            Font font = new Font();
            Assert.Equal(Font.DEFAULT_VERTICAL_ALIGN, font.VerticalAlign); // default is none
            font.VerticalAlign = value;
            Assert.Equal(value, font.VerticalAlign);
        }

        [Fact(DisplayName = "Test of the get function of the IsDefaultFont property")]
        public void IsDefaultFontTest()
        {
            Font font = new Font();
            Assert.True(font.IsDefaultFont);
            font.Italic = true;
            font.Name = "XYZ";
            Assert.False(font.IsDefaultFont);
        }

        [Fact(DisplayName = "Test of the CopyFont function")]
        public void CopyFontTest()
        {
            Font copy = exampleStyle.CopyFont();
            Assert.Equal(exampleStyle.GetHashCode(), copy.GetHashCode());
        }

        [Fact(DisplayName = "Test of the Equals method")]
        public void EqualsTest()
        {
            Font style2 = (Font)exampleStyle.Copy();
            Assert.True(exampleStyle.Equals(style2));
        }

        [Fact(DisplayName = "Test of the Equals method (inequality of Bold)")]
        public void EqualsTest2a()
        {
            Font style2 = (Font)exampleStyle.Copy();
            style2.Bold = false;
            Assert.False(exampleStyle.Equals(style2));
        }

        [Fact(DisplayName = "Test of the Equals method (inequality of Italic)")]
        public void EqualsTest2b()
        {
            Font style2 = (Font)exampleStyle.Copy();
            style2.Italic = false;
            Assert.False(exampleStyle.Equals(style2));
        }

        [Fact(DisplayName = "Test of the Equals method (inequality of Underline)")]
        public void EqualsTest2c()
        {
            Font style2 = (Font)exampleStyle.Copy();
            style2.Underline = false;
            Assert.False(exampleStyle.Equals(style2));
        }

        [Fact(DisplayName = "Test of the Equals method (inequality of DoubleUnderline)")]
        public void EqualsTest2d()
        {
            Font style2 = (Font)exampleStyle.Copy();
            style2.DoubleUnderline = false;
            Assert.False(exampleStyle.Equals(style2));
        }

        [Fact(DisplayName = "Test of the Equals method (inequality of Strike)")]
        public void EqualsTest2e()
        {
            Font style2 = (Font)exampleStyle.Copy();
            style2.Strike = false;
            Assert.False(exampleStyle.Equals(style2));
        }

        [Fact(DisplayName = "Test of the Equals method (inequality of Charset)")]
        public void EqualsTest2f()
        {
            Font style2 = (Font)exampleStyle.Copy();
            style2.Charset = "XYZ";
            Assert.False(exampleStyle.Equals(style2));
        }

        [Fact(DisplayName = "Test of the Equals method (inequality of Size)")]
        public void EqualsTest2g()
        {
            Font style2 = (Font)exampleStyle.Copy();
            style2.Size = 33;
            Assert.False(exampleStyle.Equals(style2));
        }

        [Fact(DisplayName = "Test of the Equals method (inequality of Name)")]
        public void EqualsTest2h()
        {
            Font style2 = (Font)exampleStyle.Copy();
            style2.Name = "Comic Sans";
            Assert.False(exampleStyle.Equals(style2));
        }

        [Fact(DisplayName = "Test of the Equals method (inequality of Family)")]
        public void EqualsTest2i()
        {
            Font style2 = (Font)exampleStyle.Copy();
            style2.Family = "999";
            Assert.False(exampleStyle.Equals(style2));
        }

        [Fact(DisplayName = "Test of the Equals method (inequality of ColorTheme)")]
        public void EqualsTest2j()
        {
            Font style2 = (Font)exampleStyle.Copy();
            style2.ColorTheme = 22;
            Assert.False(exampleStyle.Equals(style2));
        }

        [Fact(DisplayName = "Test of the Equals method (inequality of ColorValue)")]
        public void EqualsTest2k()
        {
            Font style2 = (Font)exampleStyle.Copy();
            style2.ColorValue = "FF9988AA";
            Assert.False(exampleStyle.Equals(style2));
        }

        [Fact(DisplayName = "Test of the Equals method (inequality of Scheme)")]
        public void EqualsTest2l()
        {
            Font style2 = (Font)exampleStyle.Copy();
            style2.Scheme = Font.SchemeValue.none;
            Assert.False(exampleStyle.Equals(style2));
        }

        [Fact(DisplayName = "Test of the Equals method (inequality of VerticalAlign)")]
        public void EqualsTest2m()
        {
            Font style2 = (Font)exampleStyle.Copy();
            style2.VerticalAlign = Font.VerticalAlignValue.none;
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
            Font copy = (Font)exampleStyle.Copy();
            Assert.False(copy.Equals(origin));
        }

        [Fact(DisplayName = "Test of the GetHashCode method (equality of two identical objects)")]
        public void GetHashCodeTest()
        {
            Font copy = (Font)exampleStyle.Copy();
            copy.InternalID = 99;  // Should not influence
            Assert.Equal(exampleStyle.GetHashCode(), copy.GetHashCode());
        }

        [Fact(DisplayName = "Test of the GetHashCode method (inequality of two different objects)")]
        public void GetHashCodeTest2()
        {
            Font copy = (Font)exampleStyle.Copy();
            copy.Bold = false;
            Assert.NotEqual(exampleStyle.GetHashCode(), copy.GetHashCode());
        }

        [Fact(DisplayName = "Test of the CompareTo method")]
        public void CompareToTest()
        {
            Font font = new Font();
            Font other = new Font();
            font.InternalID = null;
            other.InternalID = null;
            Assert.Equal(-1, font.CompareTo(other));
            font.InternalID = 5;
            Assert.Equal(1, font.CompareTo(other));
            Assert.Equal(1, font.CompareTo(null));
            other.InternalID = 5;
            Assert.Equal(0, font.CompareTo(other));
            other.InternalID = 4;
            Assert.Equal(1, font.CompareTo(other));
            other.InternalID = 6;
            Assert.Equal(-1, font.CompareTo(other));
        }

        // For code coverage
        [Fact(DisplayName = "Test of the ToString function")]
        public void ToStringTest()
        {
            Font font = new Font();
            string s1 = font.ToString();
            font.Name = "YXZ";
            Assert.NotEqual(s1, font.ToString()); // An explicit value comparison is probably not sensible
        }


    }
}
