using System.IO;
using NanoXLSX.Test.Writer_Reader.Utils;
using NanoXLSX.Themes;
using Xunit;

namespace NanoXLSX.Test.Writer_Reader.ThemesTest
{
    public class ThemeWriteReadTest
    {

        [Theory(DisplayName = "Test of the get and set function of the Name property when saving and loading a workbook")]
        [InlineData("XYZ", "XYZ")]
        [InlineData(" ", " ")]
        [InlineData("", "")]
        [InlineData(null, "")]
        public void NameTest(string name, string expectedName)
        {
            Theme theme = new Theme(0, name);
            Workbook workbook = new Workbook();
            workbook.WorkbookTheme = theme;
            Workbook givenWorkbook = TestUtils.WriteAndReadWorkbook(workbook);

            Assert.Equal(expectedName, givenWorkbook.WorkbookTheme.Name);
        }

        [Fact(DisplayName = "Test of the get and set function of the Colors property when saving and loading a workbook")]
        public void ColorsTest()
        {
            Theme theme = new Theme(0, "test");
            ColorScheme scheme = new ColorScheme();
            scheme.Name = "scheme1";
            scheme.Light2 = new SystemColor(SystemColor.Value.ButtonFace);
            scheme.Dark1 = new SrgbColor("ABCD01");
            theme.Colors = scheme;
            Workbook workbook = new Workbook();
            workbook.WorkbookTheme = theme;
            Workbook givenWorkbook = TestUtils.WriteAndReadWorkbook(workbook);

            Assert.Equal("scheme1", givenWorkbook.WorkbookTheme.Colors.Name);
            Assert.Equal(new SystemColor(SystemColor.Value.ButtonFace), givenWorkbook.WorkbookTheme.Colors.Light2);
            Assert.Equal(new SrgbColor("ABCD01"), givenWorkbook.WorkbookTheme.Colors.Dark1);
        }

        [Fact(DisplayName = "Test of the reader handling when a theme color is malformed/unknown")]
        public void UnknownThemeColorTest()
        {
            Stream stream = TestUtils.GetResource("malformed_theme_color.xlsx");
            Workbook wb = WorkbookReader.Load(stream);
            //dk1 is declared with unknown attributes
            Assert.NotNull(wb);
            Assert.Null(wb.WorkbookTheme.Colors.Dark1);
        }

    }
}
