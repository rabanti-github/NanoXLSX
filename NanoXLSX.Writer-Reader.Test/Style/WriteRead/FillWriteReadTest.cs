using NanoXLSX.Styles;
using NanoXLSX.Test.Writer_Reader.Utils;
using Xunit;
using static NanoXLSX.Styles.Fill;

namespace NanoXLSX.Test.Writer_Reader.StyleTest
{
    public class FillWriteReadTest
    {
        [Theory(DisplayName = "Test of the 'foreground' value when writing and reading a Fill style")]
        [InlineData("FFAACC00", "test")]
        [InlineData("FFAADD00", 0.5f)]
        [InlineData("FFDDCC00", true)]
        [InlineData("FFAACCDD", null)]
        public void ForegroundColorTest(string color, object value)
        {
            Style style = new Style();
            style.CurrentFill.ForegroundColor = color;
            Cell cell = TestUtils.SaveAndReadStyledCell(value, style, "A1");

            Assert.Equal(color, cell.CellStyle.CurrentFill.ForegroundColor);
            Assert.NotEqual(PatternValue.None, cell.CellStyle.CurrentFill.PatternFill);
        }

        [Theory(DisplayName = "Test of the 'background' value when writing and reading a Fill style")]
        [InlineData("FFAACC00", "test")]
        [InlineData("FFAADD00", 0.5f)]
        [InlineData("FFDDCC00", true)]
        [InlineData("FFAACCDD", null)]
        public void BackgroundColorTest(string color, object value)
        {
            Style style = new Style();
            style.CurrentFill.BackgroundColor = color;
            style.CurrentFill.PatternFill = PatternValue.DarkGray;
            Cell cell = TestUtils.SaveAndReadStyledCell(value, style, "A1");

            Assert.Equal(color, cell.CellStyle.CurrentFill.BackgroundColor);
            Assert.Equal(PatternValue.DarkGray, cell.CellStyle.CurrentFill.PatternFill);
        }

        [Theory(DisplayName = "Test of the 'patternFill' value when writing and reading a Fill style")]
        [InlineData(PatternValue.Solid, "test")]
        [InlineData(PatternValue.DarkGray, 0.5f)]
        [InlineData(PatternValue.Gray0625, true)]
        [InlineData(PatternValue.Gray125, null)]
        [InlineData(PatternValue.LightGray, "")]
        [InlineData(PatternValue.MediumGray, 0)]
        [InlineData(PatternValue.None, true)]
        public void PatternValueTest(PatternValue pattern, object value)
        {
            Style style = new Style();
            style.CurrentFill.PatternFill = pattern;
            Cell cell = TestUtils.SaveAndReadStyledCell(value, style, "A1");

            Assert.Equal(pattern, cell.CellStyle.CurrentFill.PatternFill);
        }

    }
}
