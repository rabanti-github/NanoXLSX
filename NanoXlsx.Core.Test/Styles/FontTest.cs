using System;
using NanoXLSX.Colors;
using NanoXLSX.Exceptions;
using NanoXLSX.Styles;
using NanoXLSX.Test.Core.Utils;
using Xunit;
using static NanoXLSX.Styles.Font;

namespace NanoXLSX.Test.Core.StyleTest
{
    // Ensure that these tests are executed sequentially, since static repository methods may be called 
    [Collection(nameof(SequentialCollection))]
    public class FontTest
    {

        private readonly Font exampleStyle;

        public FontTest()
        {
            exampleStyle = new Font
            {
                Bold = true,
                Italic = true,
                Underline = UnderlineValue.Double,
                Strike = true,
                Shadow = true,
                Extend = true,
                Charset = CharsetValue.ANSI,
                Size = 15,
                Name = "Arial",
                Family = FontFamilyValue.Script,
                ColorValue = "FF22AACC",
                Scheme = SchemeValue.Minor,
                VerticalAlign = VerticalTextAlignValue.Subscript
            };
        }


        [Fact(DisplayName = "Test of the default values")]
        public void DefaultValuesTest()
        {
            Assert.Equal(11f, Font.DefaultFontSize);
            Assert.Equal(FontFamilyValue.Swiss, Font.DefaultFontFamily);
            Assert.Equal(SchemeValue.Minor, Font.DefaultFontScheme);
            Assert.Equal(VerticalTextAlignValue.None, Font.DefaultVerticalAlign);
            Assert.Equal("Calibri", Font.DefaultFontName);
        }


        [Fact(DisplayName = "Test of the constructor")]
        public void ConstructorTest()
        {
            Font font = new Font();
            Assert.Equal(Font.DefaultFontSize, font.Size);
            Assert.Equal(Font.DefaultFontName, font.Name);
            Assert.Equal(Font.DefaultFontFamily, font.Family);
            Assert.Equal(Font.DefaultFontScheme, font.Scheme);
            Assert.Equal(Font.DefaultVerticalAlign, font.VerticalAlign);
            Assert.False(font.ColorValue.IsDefined);
            Assert.Null(font.ColorValue.Value);
            Assert.Equal(CharsetValue.Default, font.Charset);
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
        [InlineData(UnderlineValue.None)]
        [InlineData(UnderlineValue.DoubleAccounting)]
        [InlineData(UnderlineValue.SingleAccounting)]
        [InlineData(UnderlineValue.Double)]
        [InlineData(UnderlineValue.Single)]
        public void UnderlineTest(UnderlineValue value)
        {
            Font font = new Font();
            Assert.Equal(UnderlineValue.None, font.Underline);
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



        [Theory(DisplayName = "Test of the get and set function of the Outline property")]
        [InlineData(true)]
        [InlineData(false)]
        public void OutlineTest(bool value)
        {
            Font font = new Font();
            Assert.False(font.Outline);
            font.Outline = value;
            Assert.Equal(value, font.Outline);
        }

        [Theory(DisplayName = "Test of the get and set function of the Shadow property")]
        [InlineData(true)]
        [InlineData(false)]
        public void ShadowTest(bool value)
        {
            Font font = new Font();
            Assert.False(font.Shadow);
            font.Shadow = value;
            Assert.Equal(value, font.Shadow);
        }

        [Theory(DisplayName = "Test of the get and set function of the Condense property")]
        [InlineData(true)]
        [InlineData(false)]
        public void CondenseTest(bool value)
        {
            Font font = new Font();
            Assert.False(font.Condense);
            font.Condense = value;
            Assert.Equal(value, font.Condense);
        }

        [Theory(DisplayName = "Test of the get and set function of the Extend property")]
        [InlineData(true)]
        [InlineData(false)]
        public void ExtendTest(bool value)
        {
            Font font = new Font();
            Assert.False(font.Extend);
            font.Extend = value;
            Assert.Equal(value, font.Extend);
        }




        [Theory(DisplayName = "Test of the get and set function of the Charset property")]
        [InlineData(CharsetValue.ANSI)]
        [InlineData(CharsetValue.ApplicationDefined)]
        [InlineData(CharsetValue.Arabic)]
        [InlineData(CharsetValue.Baltic)]
        [InlineData(CharsetValue.Big5)]
        [InlineData(CharsetValue.Default)]
        [InlineData(CharsetValue.EasternEuropean)]
        [InlineData(CharsetValue.GBK)]
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
            Assert.Equal(Font.DefaultFontSize, font.Size); // 11 is default
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
            Font font = new Font
            {
                Size = givenValue
            };
            Assert.Equal(expectedValue, font.Size);
        }

        [Theory(DisplayName = "Test of the get and set function of the Name property")]
        [InlineData("Calibri")]
        [InlineData("Arial")]
        [InlineData("---")] // Not a font but a valid string
        public void NameTest(string value)
        {
            Font font = new Font();
            Assert.Equal(Font.DefaultFontName, font.Name); // Default is 'Calibri'
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
            Assert.Equal(Font.DefaultFontFamily, font.Family);
            font.Family = value;
            Assert.Equal(value, font.Family);
        }

        [Fact(DisplayName = "Test of the get and set function of the ColorValue property on sRGB (ARGB)")]
        public void ColorValueTest()
        {
            Font font = new Font();
            Assert.Equal(Color.ColorType.None, font.ColorValue.Type); // default is none
            font.ColorValue = "FFAA3322"; // implicit
            Assert.Equal(Color.ColorType.Rgb, font.ColorValue.Type);
            Assert.Equal("FFAA3322", font.ColorValue.RgbColor.ColorValue);

            Font font2 = new Font();
            Assert.Equal(Color.ColorType.None, font2.ColorValue.Type); // default is none
            font2.ColorValue = Color.CreateRgb("FFAA33AA");
            Assert.Equal(Color.ColorType.Rgb, font2.ColorValue.Type);
            Assert.Equal("FFAA33AA", font2.ColorValue.RgbColor.ColorValue);
        }

        [Fact(DisplayName = "Test of the get and set function of the ColorValue property on Indexed colors")]
        public void ColorValueTest2()
        {
            Font font = new Font();
            Assert.Equal(Color.ColorType.None, font.ColorValue.Type); // default is none
            font.ColorValue = 32; // implicit
            Assert.Equal(Color.ColorType.Indexed, font.ColorValue.Type);
            Assert.Equal(IndexedColor.Value.Navy, font.ColorValue.IndexedColor.ColorValue);

            Font font2 = new Font();
            Assert.Equal(Color.ColorType.None, font2.ColorValue.Type); // default is none
            font2.ColorValue = Color.CreateIndexed(IndexedColor.Value.Navy);
            Assert.Equal(Color.ColorType.Indexed, font.ColorValue.Type);
            Assert.Equal(IndexedColor.Value.Navy, font.ColorValue.IndexedColor.ColorValue);
        }

        [Theory(DisplayName = "Test of the failing implicit set function of the ColorValue property (invalid string values)")]
        [InlineData("77BB0")]
        [InlineData("0002200000")]
        [InlineData("XXXXXXXX")]
        public void ColorValueFailTest(string value)
        {
            Font font = new Font();
            Exception ex = Assert.Throws<StyleException>(() => font.ColorValue = value);
            Assert.Equal(typeof(StyleException), ex.GetType());
        }

        [Theory(DisplayName = "Test of the failing implicit set function of the ColorValue property (invalid int values)")]
        [InlineData(-10)]
        [InlineData(66)]
        [InlineData(999)]
        public void ColorValueFailTest2(int index)
        {
            Font font = new Font();
            Exception ex = Assert.Throws<StyleException>(() => font.ColorValue = index);
            Assert.Equal(typeof(StyleException), ex.GetType());
        }

        [Theory(DisplayName = "Test of the get and set function of the Scheme property")]
        [InlineData(SchemeValue.Major)]
        [InlineData(SchemeValue.Minor)]
        [InlineData(SchemeValue.None)]
        public void SchemeTest(SchemeValue value)
        {
            Font font = new Font();
            Assert.Equal(Font.DefaultFontScheme, font.Scheme); // default is minor
            font.Scheme = value;
            Assert.Equal(value, font.Scheme);
        }

        [Theory(DisplayName = "Test of the get and set function of the VerticalAlign property")]
        [InlineData(VerticalTextAlignValue.None)]
        [InlineData(VerticalTextAlignValue.Subscript)]
        [InlineData(VerticalTextAlignValue.Superscript)]
        [InlineData(VerticalTextAlignValue.Baseline)]
        public void VerticalAlignTest(VerticalTextAlignValue value)
        {
            Font font = new Font();
            Assert.Equal(Font.DefaultVerticalAlign, font.VerticalAlign); // default is none
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
        [InlineData("Calibri", SchemeValue.Minor)]
        [InlineData("Calibri Light", SchemeValue.Major)]
        [InlineData("Arial", SchemeValue.None)]
        [InlineData("---", SchemeValue.None)] // Not a font but a valid string
        public void ValidateFontSchemeTest(string fontName, SchemeValue scheme)
        {
            Font font = new Font
            {
                Name = fontName
            };
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
            style2.Underline = UnderlineValue.DoubleAccounting;
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
            style2.Scheme = SchemeValue.None;
            Assert.False(exampleStyle.Equals(style2));
        }

        [Fact(DisplayName = "Test of the Equals method (inequality of VerticalAlign)")]
        public void EqualsTest2m()
        {
            Font style2 = (Font)exampleStyle.Copy();
            style2.VerticalAlign = VerticalTextAlignValue.None;
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
