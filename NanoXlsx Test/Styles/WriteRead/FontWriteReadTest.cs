﻿using NanoXLSX;
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
        [InlineData(Font.UnderlineValue.u_single, "test")]
        [InlineData(Font.UnderlineValue.u_double, 0.5f)]
        [InlineData(Font.UnderlineValue.doubleAccounting, true)]
        [InlineData(Font.UnderlineValue.singleAccounting, 42)]
        [InlineData(Font.UnderlineValue.none, "")]
        public void UnderlineFontTest(Font.UnderlineValue styleValue, object value)
        {
            Style style = new Style();
            style.CurrentFont.Underline = styleValue;
            Cell cell = TestUtils.SaveAndReadStyledCell(value, style, "A1");
            Assert.Equal(styleValue, cell.CellStyle.CurrentFont.Underline);
        }

        [Theory(DisplayName = "Test of the 'vertical alignment' value when writing and reading a Font style")]
        [InlineData(Font.VerticalAlignValue.subscript, "test")]
        [InlineData(Font.VerticalAlignValue.superscript, 0.5f)]
        [InlineData(Font.VerticalAlignValue.none, true)]
        public void VerticalAlignFontTest(Font.VerticalAlignValue styleValue, object value)
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
            Style style = new Style();
            style.CurrentFont.ColorTheme = styleValue;
            Cell cell = TestUtils.SaveAndReadStyledCell(value, style, "A1");
            Assert.Equal(styleValue, cell.CellStyle.CurrentFont.ColorTheme);
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
        [InlineData(" ", "test")]
        [InlineData("test", 0.5f)]
        [InlineData("", true)]
        public void FamilyFontTest(string styleValue, object value)
        {
            Style style = new Style();
            style.CurrentFont.Family = styleValue;
            Cell cell = TestUtils.SaveAndReadStyledCell(value, style, "A1");
            Assert.Equal(styleValue, cell.CellStyle.CurrentFont.Family);
        }


        [Theory(DisplayName = "Test of the 'scheme' value when writing and reading a Font style")]
        [InlineData(Font.SchemeValue.minor, "test")]
        [InlineData(Font.SchemeValue.major, 0.5f)]
        public void SchemeFontTest(Font.SchemeValue styleValue, object value)
        {
            Style style = new Style();
            style.CurrentFont.Scheme = styleValue;
            Cell cell = TestUtils.SaveAndReadStyledCell(value, style, "A1");
            Assert.Equal(styleValue, cell.CellStyle.CurrentFont.Scheme);
        }

        [Theory(DisplayName = "Test of the 'charset' value when writing and reading a Font style")]
        [InlineData(" ", "test")]
        [InlineData("test", 0.5f)]
        [InlineData("", true)]
        public void CharsetFontTest(string styleValue, object value)
        {
            Style style = new Style();
            style.CurrentFont.Charset = styleValue;
            Cell cell = TestUtils.SaveAndReadStyledCell(value, style, "A1");
            Assert.Equal(styleValue, cell.CellStyle.CurrentFont.Charset);
        }

    }
}
