using System;
using NanoXLSX.Exceptions;
using NanoXLSX.Styles;
using Xunit;
using static NanoXLSX.Styles.Font;
using static NanoXLSX.Themes.Theme;

namespace NanoXLSX.Test.Core.StyleTest
{
    // Ensure that these tests are executed sequentially, since static repository methods may be called 
    [Collection(nameof(SequentialCollection))]
    public class FontTest
    {

        private readonly Font exampleStyle;

        public FontTest()
        {
            exampleStyle = new Font();
            exampleStyle.Bold = true;
            exampleStyle.Italic = true;
            exampleStyle.Underline = UnderlineValue.u_double;
            exampleStyle.Strike = true;
            exampleStyle.Charset = CharsetValue.ANSI;
            exampleStyle.Size = 15;
            exampleStyle.Name = "Arial";
            exampleStyle.Family = FontFamilyValue.Script;
            exampleStyle.ColorTheme = ColorSchemeElement.accent5;
            exampleStyle.ColorValue = "FF22AACC";
            exampleStyle.Scheme = SchemeValue.minor;
            exampleStyle.VerticalAlign = VerticalTextAlignValue.subscript;
        }


        [Fact(DisplayName = "Test of the default values")]
        public void DefaultValuesTest()
        {
            Assert.Equal(11f, Font.DEFAULT_FONT_SIZE);
            Assert.Equal(FontFamilyValue.Swiss, Font.DEFAULT_FONT_FAMILY);
            Assert.Equal(SchemeValue.minor, Font.DEFAULT_FONT_SCHEME);
            Assert.Equal(VerticalTextAlignValue.none, Font.DEFAULT_VERTICAL_ALIGN);
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
            Assert.Equal(CharsetValue.Default, font.Charset);
            Assert.Equal(ColorSchemeElement.light1, font.ColorTheme);
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
        [InlineData(UnderlineValue.none)]
        [InlineData(UnderlineValue.doubleAccounting)]
        [InlineData(UnderlineValue.singleAccounting)]
        [InlineData(UnderlineValue.u_double)]
        [InlineData(UnderlineValue.u_single)]
        public void UnderlineTest(UnderlineValue value)
        {
            Font font = new Font();
            Assert.Equal(UnderlineValue.none, font.Underline);
            font.Underline = value;
            Assert.Equal(value, font.Underline);
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
        [InlineData(CharsetValue.ANSI)]
        [InlineData(CharsetValue.ApplicationDefined)]
        [InlineData(CharsetValue.Arabic)]
        [InlineData(CharsetValue.Baltic)]
        [InlineData(CharsetValue.Big5)]
        [InlineData(CharsetValue.Default)]
        [InlineData(CharsetValue.EasternEuropean)]
        [InlineData(CharsetValue.GKB)]
        [InlineData(CharsetValue.Greek)]
        [InlineData(CharsetValue.Hangul)]
        [InlineData(CharsetValue.Hebrew)]
        [InlineData(CharsetValue.JIS)]
        [InlineData(CharsetValue.Johab)]
        [InlineData(CharsetValue.Macintosh)]
        [InlineData(CharsetValue.OEM)]
        [InlineData(CharsetValue.Russian)]
        [InlineData(CharsetValue.Symbols)]
        [InlineData(CharsetValue.Thai)]
        [InlineData(CharsetValue.Turkish)]
        [InlineData(CharsetValue.Vietnamese)]
        public void CharsetTest(CharsetValue value)
        {
            Font font = new Font();
            Assert.Equal(CharsetValue.Default, font.Charset);
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
        [InlineData(FontFamilyValue.NotApplicable)]
        [InlineData(FontFamilyValue.Roman)]
        [InlineData(FontFamilyValue.Swiss)]
        [InlineData(FontFamilyValue.Modern)]
        [InlineData(FontFamilyValue.Script)]
        [InlineData(FontFamilyValue.Decorative)]
        [InlineData(FontFamilyValue.Reserved1)]
        [InlineData(FontFamilyValue.Reserved2)]
        [InlineData(FontFamilyValue.Reserved3)]
        [InlineData(FontFamilyValue.Reserved4)]
        [InlineData(FontFamilyValue.Reserved5)]
        [InlineData(FontFamilyValue.Reserved6)]
        [InlineData(FontFamilyValue.Reserved7)]
        public void FamilyTest(FontFamilyValue value)
        {
            Font font = new Font();
            Assert.Equal(Font.DEFAULT_FONT_FAMILY, font.Family);
            font.Family = value;
            Assert.Equal(value, font.Family);
        }

        [Theory(DisplayName = "Test of the get and set function of the ColorTheme property")]
        [InlineData(ColorSchemeElement.dark1)]
        [InlineData(ColorSchemeElement.light1)]
        [InlineData(ColorSchemeElement.dark2)]
        [InlineData(ColorSchemeElement.light2)]
        [InlineData(ColorSchemeElement.accent1)]
        [InlineData(ColorSchemeElement.accent2)]
        [InlineData(ColorSchemeElement.accent3)]
        [InlineData(ColorSchemeElement.accent4)]
        [InlineData(ColorSchemeElement.accent5)]
        [InlineData(ColorSchemeElement.accent6)]
        [InlineData(ColorSchemeElement.hyperlink)]
        [InlineData(ColorSchemeElement.followedHyperlink)]
        public void ColorThemeTest(ColorSchemeElement element)
        {
            Font font = new Font();
            Assert.Equal(ColorSchemeElement.light1, font.ColorTheme); // light1 is default
            font.ColorTheme = element;
            Assert.Equal(element, font.ColorTheme);
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
        [InlineData(SchemeValue.major)]
        [InlineData(SchemeValue.minor)]
        [InlineData(SchemeValue.none)]
        public void SchmeTest(SchemeValue value)
        {
            Font font = new Font();
            Assert.Equal(Font.DEFAULT_FONT_SCHEME, font.Scheme); // default is minor
            font.Scheme = value;
            Assert.Equal(value, font.Scheme);
        }

        [Theory(DisplayName = "Test of the get and set function of the VerticalAlign property")]
        [InlineData(VerticalTextAlignValue.none)]
        [InlineData(VerticalTextAlignValue.subscript)]
        [InlineData(VerticalTextAlignValue.superscript)]
        public void VerticalAlignTest(VerticalTextAlignValue value)
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

        [Theory(DisplayName = "Test of the automatic assignment of font schemes on font names")]
        [InlineData("Calibri", SchemeValue.minor)]
        [InlineData("Calibri Light", SchemeValue.major)]
        [InlineData("Arial", SchemeValue.none)]
        [InlineData("---", SchemeValue.none)] // Not a font but a valid string
        public void ValidateFontSchemeTest(string fontName, SchemeValue scheme)
        {
            Font font = new Font();
            font.Name = fontName;
            Assert.Equal(scheme, font.Scheme);
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
            style2.Underline = UnderlineValue.doubleAccounting;
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
            style2.Charset = CharsetValue.Big5;
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
            style2.Family = FontFamilyValue.Reserved5;
            Assert.False(exampleStyle.Equals(style2));
        }

        [Fact(DisplayName = "Test of the Equals method (inequality of ColorTheme)")]
        public void EqualsTest2j()
        {
            Font style2 = (Font)exampleStyle.Copy();
            style2.ColorTheme = ColorSchemeElement.light2;
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
            style2.Scheme = SchemeValue.none;
            Assert.False(exampleStyle.Equals(style2));
        }

        [Fact(DisplayName = "Test of the Equals method (inequality of VerticalAlign)")]
        public void EqualsTest2m()
        {
            Font style2 = (Font)exampleStyle.Copy();
            style2.VerticalAlign = VerticalTextAlignValue.none;
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

        private static object SequentialCollection()
        {
            throw new NotImplementedException();
        }

    }
}
