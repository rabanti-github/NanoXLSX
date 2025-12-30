using NanoXLSX.Colors;
using NanoXLSX.Styles;
using NanoXLSX.Test.Writer_Reader.Utils;
using NanoXLSX.Themes;
using Xunit;
using static NanoXLSX.Styles.Fill;
using static NanoXLSX.Themes.Theme;

namespace NanoXLSX.Test.Writer_Reader.Styles
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
            var style = new Style();
            style.CurrentFill.ForegroundColor = color;
            var cell = TestUtils.SaveAndReadStyledCell(value, style, "A1");

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
            var style = new Style();
            style.CurrentFill.BackgroundColor = color;
            style.CurrentFill.PatternFill = PatternValue.DarkGray;
            var cell = TestUtils.SaveAndReadStyledCell(value, style, "A1");

            Assert.Equal(color, cell.CellStyle.CurrentFill.BackgroundColor);
            Assert.Equal(PatternValue.DarkGray, cell.CellStyle.CurrentFill.PatternFill);
        }

        [Theory(DisplayName = "Test of the 'Theme foreground color' when writing and reading a Fill style")]
        [InlineData(Theme.ColorSchemeElement.Accent1, "test")]
        [InlineData(Theme.ColorSchemeElement.Dark1, 0.5f)]
        [InlineData(Theme.ColorSchemeElement.Light1, true)]
        public void ThemeForegroundColorTest(Theme.ColorSchemeElement themeColor, object value)
        {
            var style = new Style();
            style.CurrentFill.ForegroundColor = Color.CreateTheme(themeColor);
            style.CurrentFill.PatternFill = PatternValue.Solid;

            var cell = TestUtils.SaveAndReadStyledCell(value, style, "A1");

            Assert.Equal(themeColor, cell.CellStyle.CurrentFill.ForegroundColor.ThemeColor.ColorValue);
            Assert.Null(cell.CellStyle.CurrentFill.ForegroundColor.Tint);
        }

        [Theory(DisplayName = "Test of the 'Theme foreground color with tint' when writing and reading a Fill style")]
        [InlineData(Theme.ColorSchemeElement.Accent2, -0.25)]
        [InlineData(Theme.ColorSchemeElement.Accent3, 0.5)]
        public void ThemeForegroundColorWithTintTest(Theme.ColorSchemeElement themeColor, double tint)
        {
            var style = new Style();
            style.CurrentFill.ForegroundColor = Color.CreateTheme(themeColor, tint);
            style.CurrentFill.PatternFill = PatternValue.Solid;

            var cell = TestUtils.SaveAndReadStyledCell("test", style, "A1");

            Assert.Equal(themeColor, cell.CellStyle.CurrentFill.ForegroundColor.ThemeColor.ColorValue);
            Assert.Equal(tint, cell.CellStyle.CurrentFill.ForegroundColor.Tint);
        }

        [Theory(DisplayName = "Test of the 'System foreground color' when writing and reading a Fill style")]
        [InlineData(SystemColor.Value.Menu)]
        [InlineData(SystemColor.Value.CaptionText)]
        [InlineData(SystemColor.Value.AppWorkspace)]
        public void SystemForegroundColorTest(SystemColor.Value systemColor)
        {
            var style = new Style();
            style.CurrentFill.ForegroundColor = Color.CreateSystem(systemColor);
            style.CurrentFill.PatternFill = PatternValue.Solid;

            var cell = TestUtils.SaveAndReadStyledCell("test", style, "A1");

            Assert.Equal(
                systemColor,
                cell.CellStyle.CurrentFill.ForegroundColor.SystemColor.ColorValue);
        }

        [Fact(DisplayName = "Test of the 'Auto foreground color' when writing and reading a Fill style")]
        public void AutoForegroundColorTest()
        {
            var style = new Style();
            style.CurrentFill.ForegroundColor = Color.CreateAuto();
            style.CurrentFill.PatternFill = PatternValue.Solid;

            var cell = TestUtils.SaveAndReadStyledCell("test", style, "A1");

            Assert.Equal(Color.ColorType.Auto, cell.CellStyle.CurrentFill.ForegroundColor.Type);
            Assert.True(cell.CellStyle.CurrentFill.ForegroundColor.Auto);
        }

        [Fact(DisplayName = "Test of 'None' foreground color when writing and reading a Fill style")]
        public void NoneForegroundColorTest()
        {
            var style = new Style();
            style.CurrentFill.ForegroundColor = Color.CreateNone();

            var cell = TestUtils.SaveAndReadStyledCell("test", style, "A1");

            Assert.False(cell.CellStyle.CurrentFill.ForegroundColor.IsDefined);
            Assert.Equal(Color.ColorType.None, cell.CellStyle.CurrentFill.ForegroundColor.Type);
        }

        [Fact(DisplayName = "Test implicit string to Color conversion in Fill")]
        public void ImplicitStringColorConversionTest()
        {
            var style = new Style();
            style.CurrentFill.ForegroundColor = "FF112233";

            var cell = TestUtils.SaveAndReadStyledCell("test", style, "A1");

            Assert.Equal("FF112233", cell.CellStyle.CurrentFill.ForegroundColor.RgbColor.ColorValue);
        }

        [Fact(DisplayName = "Test implicit int to IndexedColor conversion in Fill")]
        public void ImplicitIndexedColorConversionTest()
        {
            var style = new Style();
            style.CurrentFill.ForegroundColor = 10; // implicit IndexedColor
            style.CurrentFill.PatternFill = PatternValue.DarkGray;

            var cell = TestUtils.SaveAndReadStyledCell("test", style, "A1");

            Assert.Equal(10, (int)cell.CellStyle.CurrentFill.ForegroundColor.IndexedColor.ColorValue);
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
            var style = new Style();
            style.CurrentFill.PatternFill = pattern;
            var cell = TestUtils.SaveAndReadStyledCell(value, style, "A1");

            Assert.Equal(pattern, cell.CellStyle.CurrentFill.PatternFill);
        }

        [Theory(DisplayName = "Test of the 'IndexedColor' value when writing and reading a Fill style")]
        [InlineData(IndexedColor.Value.Red2, "test")]
        public void IndexedColorTest(IndexedColor.Value indexedColor, object value)
        {
            var style = new Style();
            style.CurrentFill.ForegroundColor = Color.CreateIndexed(indexedColor);
            style.CurrentFill.PatternFill = PatternValue.DarkGray; // IndexedColor requires a pattern different from 'None'
            var cell = TestUtils.SaveAndReadStyledCell(value, style, "A1");
            Assert.Equal(new IndexedColor(indexedColor), cell.CellStyle.CurrentFill.ForegroundColor.IndexedColor);

            style = new Style();
            style.CurrentFill.BackgroundColor = Color.CreateIndexed(indexedColor);
            style.CurrentFill.PatternFill = PatternValue.DarkGray; // IndexedColor requires a pattern different from 'None'
            cell = TestUtils.SaveAndReadStyledCell(value, style, "A1");

            Assert.Equal(new IndexedColor(indexedColor), cell.CellStyle.CurrentFill.BackgroundColor.IndexedColor);
        }

        [Fact(DisplayName = "Solid RGB fill writes default bgColor indexed=64 when background is not explicitly defined")]
        public void SolidRgbFillWritesDefaultIndexedBackground()
        {
            var style = new Style();
            var fill = new Fill
            {
                PatternFill = PatternValue.Solid,
                ForegroundColor = Color.CreateRgb("FF0000")
            };
            fill.BackgroundColor = Color.CreateNone();
            style.CurrentFill = fill;
            var cell = TestUtils.SaveAndReadStyledCell("test", style, "A1");

            Assert.Equal(64, (int)cell.CellStyle.CurrentFill.BackgroundColor.IndexedColor.ColorValue);
        }

        [Fact(DisplayName = "Solid Theme fill writes default bgColor indexed=64 when background is not explicitly defined")]
        public void SolidThemeFillWritesDefaultIndexedBackground()
        {
            var style = new Style();
            var fill = new Fill
            {
                PatternFill = PatternValue.Solid,
                ForegroundColor = Color.CreateTheme(ColorSchemeElement.Accent1)
            };

            fill.BackgroundColor = Color.CreateNone();
            style.CurrentFill = fill;
            var cell = TestUtils.SaveAndReadStyledCell("test", style, "A1");
            Assert.Equal(64, (int)cell.CellStyle.CurrentFill.BackgroundColor.IndexedColor.ColorValue);
        }



    }
}
