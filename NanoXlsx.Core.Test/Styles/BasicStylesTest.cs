using NanoXLSX.Shared.Exceptions;
using NanoXLSX.Shared.Exceptions;
using NanoXLSX.Styles;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xunit;
using static NanoXLSX.Shared.Enums.Styles.BorderEnums;
using static NanoXLSX.Shared.Enums.Styles.FillEnums;
using static NanoXLSX.Shared.Enums.Styles.FontEnums;
using static NanoXLSX.Shared.Enums.Styles.NumberFormatEnums;

namespace NanoXLSX_Test.Styles
{
    // Ensure that these tests are executed sequentially, since static repository methods may be called 
    [Collection(nameof(SequentialCollection))]
    public class BasicStylesTest
    {
        [Fact(DisplayName = "Test of the static Bold style")]
        public void BoldTest()
        {
            Style style = BasicStyles.Bold;
            Assert.NotNull(style);
            Assert.True(style.CurrentFont.Bold);
        }

        [Fact(DisplayName = "Test of the static Italic style")]
        public void ItalicTest()
        {
            Style style = BasicStyles.Italic;
            Assert.NotNull(style);
            Assert.True(style.CurrentFont.Italic);
        }

        [Fact(DisplayName = "Test of the static BoldItalic style")]
        public void BoldItalicTest()
        {
            Style style = BasicStyles.BoldItalic;
            Assert.NotNull(style);
            Assert.True(style.CurrentFont.Italic);
            Assert.True(style.CurrentFont.Bold);
        }

        [Fact(DisplayName = "Test of the static Underline style")]
        public void UnderlineTest()
        {
            Style style = BasicStyles.Underline;
            Assert.NotNull(style);
            Assert.Equal(UnderlineValue.u_single, style.CurrentFont.Underline);
        }

        [Fact(DisplayName = "Test of the static DoubleUnderline style")]
        public void DoubleUnderlineTest()
        {
            Style style = BasicStyles.DoubleUnderline;
            Assert.NotNull(style);
            Assert.Equal(UnderlineValue.u_double, style.CurrentFont.Underline);
        }

        [Fact(DisplayName = "Test of the static Strike style")]
        public void StrikeTest()
        {
            Style style = BasicStyles.Strike;
            Assert.NotNull(style);
            Assert.True(style.CurrentFont.Strike);
        }

        [Fact(DisplayName = "Test of the static TimeFormat style")]
        public void TimeFormatTest()
        {
            Style style = BasicStyles.TimeFormat;
            Assert.NotNull(style);
            Assert.Equal(FormatNumber.format_21, style.CurrentNumberFormat.Number);
        }

        [Fact(DisplayName = "Test of the static DateFormat style")]
        public void DateFormatTest()
        {
            Style style = BasicStyles.DateFormat;
            Assert.NotNull(style);
            Assert.Equal(FormatNumber.format_14, style.CurrentNumberFormat.Number);
        }

        [Fact(DisplayName = "Test of the static RoundFormat style")]
        public void RoundFormatTest()
        {
            Style style = BasicStyles.RoundFormat;
            Assert.NotNull(style);
            Assert.Equal(FormatNumber.format_1, style.CurrentNumberFormat.Number);
        }

        [Fact(DisplayName = "Test of the static MergeCell style")]
        public void MergeCellStyleTest()
        {
            Style style = BasicStyles.MergeCellStyle;
            Assert.NotNull(style);
            Assert.True(style.CurrentCellXf.ForceApplyAlignment);
        }

        [Fact(DisplayName = "Test of the static DottedFill_0_125 style")]
        public void DottedFill_0_125Test()
        {
            Style style = BasicStyles.DottedFill_0_125;
            Assert.NotNull(style);
            Assert.Equal(PatternValue.gray125, style.CurrentFill.PatternFill);
        }

        [Fact(DisplayName = "Test of the static BorderFrame style")]
        public void BorderFrameTest()
        {
            Style style = BasicStyles.BorderFrame;
            Assert.NotNull(style);
            Assert.Equal(StyleValue.thin, style.CurrentBorder.TopStyle);
            Assert.Equal(StyleValue.thin, style.CurrentBorder.BottomStyle);
            Assert.Equal(StyleValue.thin, style.CurrentBorder.LeftStyle);
            Assert.Equal(StyleValue.thin, style.CurrentBorder.RightStyle);
        }

        [Fact(DisplayName = "Test of the static BorderFrameHeader style")]
        public void BorderFrameHeaderTest()
        {
            Style style = BasicStyles.BorderFrameHeader;
            Assert.NotNull(style);
            Assert.Equal(StyleValue.thin, style.CurrentBorder.TopStyle);
            Assert.Equal(StyleValue.medium, style.CurrentBorder.BottomStyle);
            Assert.Equal(StyleValue.thin, style.CurrentBorder.LeftStyle);
            Assert.Equal(StyleValue.thin, style.CurrentBorder.RightStyle);
            Assert.True(style.CurrentFont.Bold);
        }

        [Theory(DisplayName = "Test of the ColorizedText function")]
        [InlineData("000000", "FF000000")]
        [InlineData("3CDEF0", "FF3CDEF0")]
        [InlineData("af3cd1", "FFAF3CD1")]
        [InlineData("FFFFFF", "FFFFFFFF")]
        public void ColorizedTextTest(string hexCode, string expectedHexCode)
        {
            Style style = BasicStyles.ColorizedText(hexCode);
            Assert.NotNull(style);
            Assert.Equal(expectedHexCode, style.CurrentFont.ColorValue);
        }

        [Theory(DisplayName = "Test of the failing ColorizedText function")]
        [InlineData(null)]
        [InlineData("")]
        [InlineData(" ")]
        [InlineData("AAFF")]
        [InlineData("AAFFCC22")]
        [InlineData("XXXXVV")]
        public void ColorizedTextFailTest(string hexCode)
        {
            Assert.Throws<StyleException>(() => BasicStyles.ColorizedText(hexCode));
        }

        [Theory(DisplayName = "Test of the ColorizedBackground function")]
        [InlineData("000000", "FF000000")]
        [InlineData("3CDEF0", "FF3CDEF0")]
        [InlineData("af3cd1", "FFAF3CD1")]
        [InlineData("FFFFFF", "FFFFFFFF")]
        public void ColorizedBackgroundTest(string hexCode, string expectedHexCode)
        {
            Style style = BasicStyles.ColorizedBackground(hexCode);
            Assert.NotNull(style);
            Assert.Equal(expectedHexCode, style.CurrentFill.ForegroundColor);
            Assert.Equal(Fill.DEFAULT_COLOR, style.CurrentFill.BackgroundColor);
            Assert.Equal(PatternValue.solid, style.CurrentFill.PatternFill);

        }

        [Theory(DisplayName = "Test of the failing ColorizedBackground function")]
        [InlineData(null)]
        [InlineData("")]
        [InlineData(" ")]
        [InlineData("AAFF")]
        [InlineData("AAFFCC22")]
        [InlineData("XXXXVV")]
        public void ColorizedBackgroundFailTest(string hexCode)
        {
            Assert.Throws<StyleException>(() => BasicStyles.ColorizedBackground(hexCode));
        }

        [Theory(DisplayName = "Test of the Font function with name")]
        [InlineData("Calibri")]
        [InlineData("Arial")]
        [InlineData("Times New Roman")]
        [InlineData("Sans Serif")]
        [InlineData("Tahoma")]
        public void FontTest(string name)
        {
            Style style = BasicStyles.Font(name);
            Assert.Equal(name, style.CurrentFont.Name);
            Assert.Equal(Font.DEFAULT_FONT_SIZE, style.CurrentFont.Size);
            Assert.False(style.CurrentFont.Bold);
            Assert.False(style.CurrentFont.Italic);
        }

        [Theory(DisplayName = "Test of the Font function with name and size")]
        [InlineData("Calibri", 12f)]
        [InlineData("Arial", 1f)]
        [InlineData("Times New Roman", 409f)]
        [InlineData("Sans Serif", 50f)]
        [InlineData("Tahoma", 11f)]
        public void FontTest2(string name, float size)
        {
            Style style = BasicStyles.Font(name, size);
            Assert.Equal(name, style.CurrentFont.Name);
            Assert.Equal(size, style.CurrentFont.Size);
            Assert.False(style.CurrentFont.Bold);
            Assert.False(style.CurrentFont.Italic);
        }

        [Theory(DisplayName = "Test of the Font function with name, size and bold state")]
        [InlineData("Calibri", 12f, false)]
        [InlineData("Arial", 1f, false)]
        [InlineData("Times New Roman", 409f, true)]
        [InlineData("Sans Serif", 50f, false)]
        [InlineData("Tahoma", 11f, true)]
        public void FontTest3(string name, float size, bool bold)
        {
            Style style = BasicStyles.Font(name, size, bold);
            Assert.Equal(name, style.CurrentFont.Name);
            Assert.Equal(size, style.CurrentFont.Size);
            Assert.Equal(bold, style.CurrentFont.Bold);
            Assert.False(style.CurrentFont.Italic);
        }

        [Fact(DisplayName = "Test of the Font function for the auto adjustment of invalid font sizes")]
        public void FontTest4()
        {
            Style style = BasicStyles.Font("Arial", -1f);
            Assert.Equal(Font.MIN_FONT_SIZE, style.CurrentFont.Size);
            style = BasicStyles.Font("Arial", 0.5f);
            Assert.Equal(Font.MIN_FONT_SIZE, style.CurrentFont.Size);
            style = BasicStyles.Font("Arial", 409.1f);
            Assert.Equal(Font.MAX_FONT_SIZE, style.CurrentFont.Size);
            style = BasicStyles.Font("Arial", 1000f);
            Assert.Equal(Font.MAX_FONT_SIZE, style.CurrentFont.Size);
        }

        [Fact(DisplayName = "Test of the failing Font function on a invalid font name")]
        public void FontFailTest()
        {
            Assert.Throws<StyleException>(() => BasicStyles.Font(null));
            Assert.Throws<StyleException>(() => BasicStyles.Font(""));
        }

        private static object SequentialCollection()
        {
            throw new NotImplementedException();
        }

    }
}
