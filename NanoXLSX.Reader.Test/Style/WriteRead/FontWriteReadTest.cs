using System;
using NanoXLSX.Schemes;
using NanoXLSX;
using NanoXLSX.Styles;
using Xunit;

namespace NanoXLSX_Test.Styles.WriteRead
{
    public class FontWriteReadTest
    {

        [Theory(DisplayName = "Test of the 'bold' value when writing and reading a Font style")]
        [InlineData(true, "test")]
        [InlineData(false, 0.5f)]
        public void BoldFontTest(bool styleValue, object value)
        {
            Style style = new Style();
            style.CurrentFont.Bold = styleValue;
            Cell cell = TestUtils.SaveAndReadStyledCell(value, style, "A1");
            Assert.Equal(styleValue, cell.CellStyle.CurrentFont.Bold);
        }

        [Theory(DisplayName = "Test of the 'italic' value when writing and reading a Font style")]
        [InlineData(true, "test")]
        [InlineData(false, 0.5f)]
        public void ItalicFontTest(bool styleValue, object value)
        {
            Style style = new Style();
            style.CurrentFont.Italic = styleValue;
            Cell cell = TestUtils.SaveAndReadStyledCell(value, style, "A1");
            Assert.Equal(styleValue, cell.CellStyle.CurrentFont.Italic);
        }

        [Theory(DisplayName = "Test of the 'strike' value when writing and reading a Font style")]
        [InlineData(true, "test")]
        [InlineData(false, 0.5f)]
        public void StrikeFontTest(bool styleValue, object value)
        {
            Style style = new Style();
            style.CurrentFont.Strike = styleValue;
            Cell cell = TestUtils.SaveAndReadStyledCell(value, style, "A1");
            Assert.Equal(styleValue, cell.CellStyle.CurrentFont.Strike);
        }

        [Theory(DisplayName = "Test of the 'underline' value when writing and reading a Font style")]
        [InlineData(UnderlineValue.u_single, "test")]
        [InlineData(UnderlineValue.u_double, 0.5f)]
        [InlineData(UnderlineValue.doubleAccounting, true)]
        [InlineData(UnderlineValue.singleAccounting, 42)]
        [InlineData(UnderlineValue.none, "")]
        public void UnderlineFontTest(UnderlineValue styleValue, object value)
        {
            Style style = new Style();
            style.CurrentFont.Underline = styleValue;
            Cell cell = TestUtils.SaveAndReadStyledCell(value, style, "A1");
            Assert.Equal(styleValue, cell.CellStyle.CurrentFont.Underline);
        }

        [Theory(DisplayName = "Test of the 'vertical alignment' value when writing and reading a Font style")]
        [InlineData(VerticalTextAlignValue.subscript, "test")]
        [InlineData(VerticalTextAlignValue.superscript, 0.5f)]
        [InlineData(VerticalTextAlignValue.none, true)]
        public void VerticalAlignFontTest(VerticalTextAlignValue styleValue, object value)
        {
            Style style = new Style();
            style.CurrentFont.VerticalAlign = styleValue;
            Cell cell = TestUtils.SaveAndReadStyledCell(value, style, "A1");
            Assert.Equal(styleValue, cell.CellStyle.CurrentFont.VerticalAlign);
        }

        [Theory(DisplayName = "Test of the 'size' value when writing and reading a Font style")]
        [InlineData(10.5f, "test")]
        [InlineData(11f, 0.5f)]
        [InlineData(50.55f, true)]
        public void SizeFontTest(float styleValue, object value)
        {
            Style style = new Style();
            style.CurrentFont.Size = styleValue;
            Cell cell = TestUtils.SaveAndReadStyledCell(value, style, "A1");
            Assert.Equal(styleValue, cell.CellStyle.CurrentFont.Size);
        }

        [Theory(DisplayName = "Test of the 'theme' value when writing and reading a Font style")]
        [InlineData(ColorSchemeElement.dark1, "test")]
        [InlineData(ColorSchemeElement.light1, 0.5f)]
        [InlineData(ColorSchemeElement.dark2, true)]
        [InlineData(ColorSchemeElement.light2, 42)]
        [InlineData(ColorSchemeElement.accent1, false)]
        [InlineData(ColorSchemeElement.accent2, null)]
        [InlineData(ColorSchemeElement.accent3, " ")]
        [InlineData(ColorSchemeElement.accent4, -3.33f)]
        [InlineData(ColorSchemeElement.accent5, 0)]
        [InlineData(ColorSchemeElement.accent6, "")]
        [InlineData(ColorSchemeElement.hyperlink, "test")]
        [InlineData(ColorSchemeElement.followedHyperlink, 0.5f)]

        public void ThemeFontTest(ColorSchemeElement element, object value)
        {
            Style style = new Style();
            style.CurrentFont.ColorTheme = element;
            Cell cell = TestUtils.SaveAndReadStyledCell(value, style, "A1");
            Assert.Equal(element, cell.CellStyle.CurrentFont.ColorTheme);
        }

        [Theory(DisplayName = "Test of the 'colorValue' value when writing and reading a Font style")]
        [InlineData("FFAABBCC", "test")]
        [InlineData("", 0.5f)]
        public void ColorValueFontTest(string styleValue, object value)
        {
            Style style = new Style();
            style.CurrentFont.ColorValue = styleValue;
            Cell cell = TestUtils.SaveAndReadStyledCell(value, style, "A1");
            Assert.Equal(styleValue, cell.CellStyle.CurrentFont.ColorValue);
        }

        [Theory(DisplayName = "Test of the 'name' value when writing and reading a Font style")]
        [InlineData(" ", "test")]
        [InlineData("test", 0.5f)]
        public void NameFontTest(string styleValue, object value)
        {
            Style style = new Style();
            style.CurrentFont.Name = styleValue;
            Cell cell = TestUtils.SaveAndReadStyledCell(value, style, "A1");
            Assert.Equal(styleValue, cell.CellStyle.CurrentFont.Name);
        }

        [Theory(DisplayName = "Test of the 'family' value when writing and reading a Font style")]
        [InlineData(FontFamilyValue.Decorative, "test", "test")]
        [InlineData(FontFamilyValue.Modern, 0.5f, 0.5f)]
        [InlineData(FontFamilyValue.Roman, true, true)]
        [InlineData(FontFamilyValue.Script, 42, 42)]
        [InlineData(FontFamilyValue.Swiss, null, null)]
        [InlineData(FontFamilyValue.NotApplicable, "", "")]
        [InlineData(FontFamilyValue.Reserved1, -0.55f, -0.55f)]
        [InlineData(FontFamilyValue.Reserved2, "test", "test")]
        [InlineData(FontFamilyValue.Reserved3, false, false)]
        [InlineData(FontFamilyValue.Reserved4, 0, 0)]
        [InlineData(FontFamilyValue.Reserved5, long.MaxValue, long.MaxValue)]
        [InlineData(FontFamilyValue.Reserved6, float.MinValue, float.MinValue)]
        [InlineData(FontFamilyValue.Reserved7, uint.MaxValue, uint.MaxValue)]
        [InlineData(FontFamilyValue.Reserved8, ulong.MaxValue, ulong.MaxValue)]
        [InlineData(FontFamilyValue.Reserved9, SByte.MaxValue, 127)]
        public void FamilyFontTest(FontFamilyValue styleValue, object givenValue, object expectdValue)
        {
            Style style = new Style();

            style.CurrentFont.Family = styleValue;
            Cell cell = TestUtils.SaveAndReadStyledCell(givenValue, expectdValue, style, "A1");
            Assert.Equal(styleValue, cell.CellStyle.CurrentFont.Family);
        }


        [Theory(DisplayName = "Test of the 'scheme' value when writing and reading a Font style")]
        [InlineData(SchemeValue.minor, "test")]
        [InlineData(SchemeValue.major, 0.5f)]
        public void SchemeFontTest(SchemeValue styleValue, object value)
        {
            Style style = new Style();
            style.CurrentFont.Scheme = styleValue;
            Cell cell = TestUtils.SaveAndReadStyledCell(value, style, "A1");
            Assert.Equal(styleValue, cell.CellStyle.CurrentFont.Scheme);
        }

        [Theory(DisplayName = "Test of the 'charset' value when writing and reading a Font style")]
        [InlineData(CharsetValue.ANSI, "test", "test")]
        [InlineData(CharsetValue.ApplicationDefined, 0.5f, 0.5f)]
        [InlineData(CharsetValue.Arabic, true, true)]
        [InlineData(CharsetValue.Baltic, 42, 42)]
        [InlineData(CharsetValue.Big5, null, null)]
        [InlineData(CharsetValue.Default, "", "")]
        [InlineData(CharsetValue.EasternEuropean, -0.55d, -0.55f)]
        [InlineData(CharsetValue.GKB, false, false)]
        [InlineData(CharsetValue.Greek, int.MaxValue, int.MaxValue)]
        [InlineData(CharsetValue.Hangul, double.MaxValue, double.MaxValue)]
        [InlineData(CharsetValue.Hebrew, float.MinValue, float.MinValue)]
        [InlineData(CharsetValue.JIS, SByte.MaxValue, 127)]
        [InlineData(CharsetValue.Johab, uint.MaxValue, uint.MaxValue)]
        [InlineData(CharsetValue.Macintosh, long.MaxValue, long.MaxValue)]
        [InlineData(CharsetValue.OEM, ulong.MaxValue, ulong.MaxValue)]
        [InlineData(CharsetValue.Russian, -1, -1)]
        [InlineData(CharsetValue.Symbols, uint.MinValue, 0)]
        [InlineData(CharsetValue.Thai, " ", " ")]
        [InlineData(CharsetValue.Turkish, 0.0f, 0)]
        [InlineData(CharsetValue.Vietnamese, 0x0, 0)]
        public void CharsetFontTest(CharsetValue styleValue, object givenValue, object expectedValue)
        {
            Style style = new Style();
            style.CurrentFont.Charset = styleValue;
            Cell cell = TestUtils.SaveAndReadStyledCell(givenValue, expectedValue, style, "A1");
            Assert.Equal(styleValue, cell.CellStyle.CurrentFont.Charset);
        }

    }
}
