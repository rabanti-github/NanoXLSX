using NanoXLSX.Themes;
using NanoXLSX_Test;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xunit;

namespace NanoXLSX.Core_Test.Themes
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
