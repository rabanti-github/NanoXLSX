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

    }
}
