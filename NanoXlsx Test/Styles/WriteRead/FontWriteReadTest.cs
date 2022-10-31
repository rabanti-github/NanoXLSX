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
        [InlineData(1, "test")]
        [InlineData(2, 0.5f)]
        [InlineData(64, true)]
        public void ThemeFontTest(int styleValue, object value)
        {
            ColorScheme scheme = new ColorScheme(styleValue);
            Style style = new Style();
            style.CurrentFont.ColorTheme = scheme;
            Cell cell = TestUtils.SaveAndReadStyledCell(value, style, "A1");
            Assert.Equal(styleValue, cell.CellStyle.CurrentFont.ColorTheme.GetSchemeId());
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
        [InlineData(FontEnums.FontFamilyValue.Decorative, "test")]
        [InlineData(FontEnums.FontFamilyValue.Modern, 0.5f)]
        [InlineData(FontEnums.FontFamilyValue.Roman, true)]
        [InlineData(FontEnums.FontFamilyValue.Script, 42)]
        [InlineData(FontEnums.FontFamilyValue.Swiss, null)]
        [InlineData(FontEnums.FontFamilyValue.NotApplicable, "")]
        [InlineData(FontEnums.FontFamilyValue.Reserved1, -0.55f)]
        [InlineData(FontEnums.FontFamilyValue.Reserved2, "test")]
        [InlineData(FontEnums.FontFamilyValue.Reserved3, false)]
        [InlineData(FontEnums.FontFamilyValue.Reserved4, 0)]
        [InlineData(FontEnums.FontFamilyValue.Reserved5, long.MaxValue)]
        [InlineData(FontEnums.FontFamilyValue.Reserved6, float.MinValue)]
        [InlineData(FontEnums.FontFamilyValue.Reserved7, uint.MaxValue)]
        [InlineData(FontEnums.FontFamilyValue.Reserved8, ulong.MaxValue)]
        [InlineData(FontEnums.FontFamilyValue.Reserved9, SByte.MaxValue)]
        public void FamilyFontTest(FontEnums.FontFamilyValue styleValue, object value)
        {
            Style style = new Style();
            
            style.CurrentFont.Family = styleValue;
            Cell cell = TestUtils.SaveAndReadStyledCell(value, style, "A1");
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
        [InlineData(FontEnums.CharsetValue.ANSI, "test")]
        [InlineData(FontEnums.CharsetValue.ApplicationDefined, 0.5f)]
        [InlineData(FontEnums.CharsetValue.Arabic, true)]
        [InlineData(FontEnums.CharsetValue.Baltic, 42)]
        [InlineData(FontEnums.CharsetValue.Big5, null)]
        [InlineData(FontEnums.CharsetValue.Default, "")]
        [InlineData(FontEnums.CharsetValue.EasternEuropean, -0.55d)]
        [InlineData(FontEnums.CharsetValue.GKB, false)]
        [InlineData(FontEnums.CharsetValue.Greek, int.MaxValue)]
        [InlineData(FontEnums.CharsetValue.Hangul, double.MaxValue)]
        [InlineData(FontEnums.CharsetValue.Hebrew, float.MinValue)]
        [InlineData(FontEnums.CharsetValue.JIS, SByte.MaxValue)]
        [InlineData(FontEnums.CharsetValue.Johab, uint.MaxValue)]
        [InlineData(FontEnums.CharsetValue.Macintosh, long.MaxValue)]
        [InlineData(FontEnums.CharsetValue.OEM, ulong.MaxValue)]
        [InlineData(FontEnums.CharsetValue.Russian, -1)]
        [InlineData(FontEnums.CharsetValue.Symbols, uint.MinValue)]
        [InlineData(FontEnums.CharsetValue.Thai, " ")]
        [InlineData(FontEnums.CharsetValue.Turkish, 0.0f)]
        [InlineData(FontEnums.CharsetValue.Vietnamese, 0x0)]
        public void CharsetFontTest(FontEnums.CharsetValue styleValue, object value)
        {
            Style style = new Style();
            style.CurrentFont.Charset = styleValue;
            Cell cell = TestUtils.SaveAndReadStyledCell(value, style, "A1");
            Assert.Equal(styleValue, cell.CellStyle.CurrentFont.Charset);
        }

    }
}
