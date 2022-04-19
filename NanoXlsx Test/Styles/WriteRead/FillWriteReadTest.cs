using NanoXLSX;
using NanoXLSX.Styles;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xunit;

namespace NanoXLSX_Test.Styles.WriteRead
{
    public class FillWriteReadTest
    {
        [Theory(DisplayName = "Test of the 'foreground' value when writing and reading a Fill style")]
        [InlineData("FFAACC00", "test")]
        [InlineData("FFAADD00", 0.5f)]
        [InlineData("FFDDCC00", true)]
        [InlineData("FFAACCDD", null)]
        public void ForegroundColorTest(String color, object value)
        {
            Style style = new Style();
            style.CurrentFill.ForegroundColor = color;
            Cell cell = TestUtils.SaveAndReadStyledCell(value, style, "A1");

            Assert.Equal(color, cell.CellStyle.CurrentFill.ForegroundColor);
            Assert.NotEqual(Fill.PatternValue.none, cell.CellStyle.CurrentFill.PatternFill);
        }

        [Theory(DisplayName = "Test of the 'background' value when writing and reading a Fill style")]
        [InlineData("FFAACC00", "test")]
        [InlineData("FFAADD00", 0.5f)]
        [InlineData("FFDDCC00", true)]
        [InlineData("FFAACCDD", null)]
        public void BackgroundColorTest(String color, object value)
        {
            Style style = new Style();
            style.CurrentFill.BackgroundColor = color;
            style.CurrentFill.PatternFill = Fill.PatternValue.darkGray;
            Cell cell = TestUtils.SaveAndReadStyledCell(value, style, "A1");

            Assert.Equal(color, cell.CellStyle.CurrentFill.BackgroundColor);
            Assert.Equal(Fill.PatternValue.darkGray, cell.CellStyle.CurrentFill.PatternFill);
        }

        [Theory(DisplayName = "Test of the 'patternFill' value when writing and reading a Fill style")]
        [InlineData(Fill.PatternValue.solid, "test")]
        [InlineData(Fill.PatternValue.darkGray, 0.5f)]
        [InlineData(Fill.PatternValue.gray0625, true)]
        [InlineData(Fill.PatternValue.gray125, null)]
        [InlineData(Fill.PatternValue.lightGray, "")]
        [InlineData(Fill.PatternValue.mediumGray, 0)]
        [InlineData(Fill.PatternValue.none, true)]
        public void PatternValueTest(Fill.PatternValue pattern, object value)
        {
            Style style = new Style();
            style.CurrentFill.PatternFill = pattern;
            Cell cell = TestUtils.SaveAndReadStyledCell(value, style, "A1");

            Assert.Equal(pattern, cell.CellStyle.CurrentFill.PatternFill);
        }

    }
}
