using NanoXLS.Shared.Enums.Schemes;
using NanoXLSX;
using NanoXLSX.Shared.Enums.Styles;
using NanoXLSX.Styles;
using NanoXLSX.Themes;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
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
        [InlineData(FontEnums.UnderlineValue.u_single, "test")]
        [InlineData(FontEnums.UnderlineValue.u_double, 0.5f)]
        [InlineData(FontEnums.UnderlineValue.doubleAccounting, true)]
        [InlineData(FontEnums.UnderlineValue.singleAccounting, 42)]
        [InlineData(FontEnums.UnderlineValue.none, "")]
        public void UnderlineFontTest(FontEnums.UnderlineValue styleValue, object value)
        {
            Style style = new Style();
            style.CurrentFont.Underline = styleValue;
            Cell cell = TestUtils.SaveAndReadStyledCell(value, style, "A1");
            Assert.Equal(styleValue, cell.CellStyle.CurrentFont.Underline);
        }

        [Theory(DisplayName = "Test of the 'vertical alignment' value when writing and reading a Font style")]
        [InlineData(FontEnums.VerticalTextAlignValue.subscript, "test")]
        [InlineData(FontEnums.VerticalTextAlignValue.superscript, 0.5f)]
        [InlineData(FontEnums.VerticalTextAlignValue.none, true)]
        public void VerticalAlignFontTest(FontEnums.VerticalTextAlignValue styleValue, object value)
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
        [InlineData(ThemeEnums.ColorSchemeElement.dark1, "test")]
        [InlineData(ThemeEnums.ColorSchemeElement.light1, 0.5f)]
        [InlineData(ThemeEnums.ColorSchemeElement.dark2, true)]
        [InlineData(ThemeEnums.ColorSchemeElement.light2, 42)]
        [InlineData(ThemeEnums.ColorSchemeElement.accent1, false)]
        [InlineData(ThemeEnums.ColorSchemeElement.accent2, null)]
        [InlineData(ThemeEnums.ColorSchemeElement.accent3, " ")]
        [InlineData(ThemeEnums.ColorSchemeElement.accent4, -3.33f)]
        [InlineData(ThemeEnums.ColorSchemeElement.accent5, 0)]
        [InlineData(ThemeEnums.ColorSchemeElement.accent6, "")]
        [InlineData(ThemeEnums.ColorSchemeElement.hyperlink, "test")]
        [InlineData(ThemeEnums.ColorSchemeElement.followedHyperlink, 0.5f)]

        public void ThemeFontTest(ThemeEnums.ColorSchemeElement element, object value)
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
        [InlineData(FontEnums.FontFamilyValue.Decorative, "test", "test")]
        [InlineData(FontEnums.FontFamilyValue.Modern, 0.5f, 0.5f)]
        [InlineData(FontEnums.FontFamilyValue.Roman, true, true)]
        [InlineData(FontEnums.FontFamilyValue.Script, 42, 42)]
        [InlineData(FontEnums.FontFamilyValue.Swiss, null, null)]
        [InlineData(FontEnums.FontFamilyValue.NotApplicable, "", "")]
        [InlineData(FontEnums.FontFamilyValue.Reserved1, -0.55f, -0.55f)]
        [InlineData(FontEnums.FontFamilyValue.Reserved2, "test", "test")]
        [InlineData(FontEnums.FontFamilyValue.Reserved3, false, false)]
        [InlineData(FontEnums.FontFamilyValue.Reserved4, 0, 0)]
        [InlineData(FontEnums.FontFamilyValue.Reserved5, long.MaxValue, long.MaxValue)]
        [InlineData(FontEnums.FontFamilyValue.Reserved6, float.MinValue, float.MinValue)]
        [InlineData(FontEnums.FontFamilyValue.Reserved7, uint.MaxValue, uint.MaxValue)]
        [InlineData(FontEnums.FontFamilyValue.Reserved8, ulong.MaxValue, ulong.MaxValue)]
        [InlineData(FontEnums.FontFamilyValue.Reserved9, SByte.MaxValue, 127)]
        public void FamilyFontTest(FontEnums.FontFamilyValue styleValue, object givenValue, object expectdValue)
        {
            Style style = new Style();
            
            style.CurrentFont.Family = styleValue;
            Cell cell = TestUtils.SaveAndReadStyledCell(givenValue, expectdValue, style, "A1");
            Assert.Equal(styleValue, cell.CellStyle.CurrentFont.Family);
        }


        [Theory(DisplayName = "Test of the 'scheme' value when writing and reading a Font style")]
        [InlineData(FontEnums.SchemeValue.minor, "test")]
        [InlineData(FontEnums.SchemeValue.major, 0.5f)]
        public void SchemeFontTest(FontEnums.SchemeValue styleValue, object value)
        {
            Style style = new Style();
            style.CurrentFont.Scheme = styleValue;
            Cell cell = TestUtils.SaveAndReadStyledCell(value, style, "A1");
            Assert.Equal(styleValue, cell.CellStyle.CurrentFont.Scheme);
        }

        [Theory(DisplayName = "Test of the 'charset' value when writing and reading a Font style")]
        [InlineData(FontEnums.CharsetValue.ANSI, "test", "test")]
        [InlineData(FontEnums.CharsetValue.ApplicationDefined, 0.5f, 0.5f)]
        [InlineData(FontEnums.CharsetValue.Arabic, true, true)]
        [InlineData(FontEnums.CharsetValue.Baltic, 42, 42)]
        [InlineData(FontEnums.CharsetValue.Big5, null, null)]
        [InlineData(FontEnums.CharsetValue.Default, "", "")]
        [InlineData(FontEnums.CharsetValue.EasternEuropean, -0.55d, -0.55f)]
        [InlineData(FontEnums.CharsetValue.GKB, false, false)]
        [InlineData(FontEnums.CharsetValue.Greek, int.MaxValue, int.MaxValue)]
        [InlineData(FontEnums.CharsetValue.Hangul, double.MaxValue, double.MaxValue)]
        [InlineData(FontEnums.CharsetValue.Hebrew, float.MinValue, float.MinValue)]
        [InlineData(FontEnums.CharsetValue.JIS, SByte.MaxValue, 127)]
        [InlineData(FontEnums.CharsetValue.Johab, uint.MaxValue, uint.MaxValue)]
        [InlineData(FontEnums.CharsetValue.Macintosh, long.MaxValue, long.MaxValue)]
        [InlineData(FontEnums.CharsetValue.OEM, ulong.MaxValue, ulong.MaxValue)]
        [InlineData(FontEnums.CharsetValue.Russian, -1, -1)]
        [InlineData(FontEnums.CharsetValue.Symbols, uint.MinValue, 0)]
        [InlineData(FontEnums.CharsetValue.Thai, " ", " ")]
        [InlineData(FontEnums.CharsetValue.Turkish, 0.0f, 0)]
        [InlineData(FontEnums.CharsetValue.Vietnamese, 0x0, 0)]
        public void CharsetFontTest(FontEnums.CharsetValue styleValue, object givenValue, object expectedValue)
        {
            Style style = new Style();
            style.CurrentFont.Charset = styleValue;
            Cell cell = TestUtils.SaveAndReadStyledCell(givenValue, expectedValue, style, "A1");
            Assert.Equal(styleValue, cell.CellStyle.CurrentFont.Charset);
        }

    }
}
