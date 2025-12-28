using NanoXLSX.Colors;
using NanoXLSX.Test.Writer_Reader.Utils;
using NanoXLSX.Themes;
using Xunit;

namespace NanoXLSX.Test.Writer_Reader.Themes
{
    public class SrgbColorWriteReadTest
    {
        [Theory(DisplayName = "Test of the correct reader parsing of SystemColors, when saving and loading a workbook")]
        [InlineData("000000", "FF000000")]
        [InlineData("111111", "FF111111")]
        [InlineData("aaaaaa", "FFAAAAAA")]
        [InlineData("aBcDeF", "FFABCDEF")]
        [InlineData("0A2B3C", "FF0A2B3C")]
        [InlineData("0a2B3c", "FF0A2B3C")]
        [InlineData("ffffff", "FFFFFFFF")]
        [InlineData("FFFFFF", "FFFFFFFF")]
        public void SystemColorReadWriteTest(string givenColor, string expectedColor)
        {
            Theme theme = new Theme("test");
            SrgbColor color = new SrgbColor(givenColor);
            theme.Colors.Dark1 = color;
            Workbook workbook = new Workbook
            {
                WorkbookTheme = theme
            };
            Assert.Equal(expectedColor, ((SrgbColor)workbook.WorkbookTheme.Colors.Dark1).ColorValue); // already UC
            Assert.Equal(expectedColor, ((SrgbColor)workbook.WorkbookTheme.Colors.Dark1).StringValue); // already UC
            Workbook givenWorkbook = TestUtils.WriteAndReadWorkbook(workbook);
            Assert.Equal(expectedColor, ((SrgbColor)givenWorkbook.WorkbookTheme.Colors.Dark1).ColorValue);
            Assert.Equal(expectedColor, ((SrgbColor)givenWorkbook.WorkbookTheme.Colors.Dark1).StringValue);
        }
    }
}
