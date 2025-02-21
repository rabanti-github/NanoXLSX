using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NanoXLSX.Test.Writer_Reader.Utils;
using NanoXLSX.Themes;
using Xunit;

namespace NanoXLSX.Test.Writer_Reader.Themes
{
    public class SystemColorWriteReadTest
    {
        [Theory(DisplayName = "Test of the correct reader parsing of SystemColors, when saving and loading a workbook")]
        [InlineData(SystemColor.Value.ThreeDimensionalDarkShadow, "111111")]
        [InlineData(SystemColor.Value.ThreeDimensionalLight, "222222")]
        [InlineData(SystemColor.Value.ActiveBorder, "FFAA05")]
        [InlineData(SystemColor.Value.ActiveCaption, "333333")]
        [InlineData(SystemColor.Value.AppWorkspace, "444444")]
        [InlineData(SystemColor.Value.Background, "555555")]
        [InlineData(SystemColor.Value.ButtonFace, "666666")]
        [InlineData(SystemColor.Value.ButtonHighlight, "777777")]
        [InlineData(SystemColor.Value.ButtonShadow, "888888")]
        [InlineData(SystemColor.Value.ButtonText, "999999")]
        [InlineData(SystemColor.Value.CaptionText, "AAAAAA")]
        [InlineData(SystemColor.Value.GradientActiveCaption, "BBBBBB")]
        [InlineData(SystemColor.Value.GradientInactiveCaption, "CCCCCC")]
        [InlineData(SystemColor.Value.GrayText, "DDDDDD")]
        [InlineData(SystemColor.Value.Highlight, "EEEEEE")]
        [InlineData(SystemColor.Value.HighlightText, "FFFFFF")]
        [InlineData(SystemColor.Value.HotLight, "ABCDEF")]
        [InlineData(SystemColor.Value.InactiveBorder, "FEDCBA")]
        [InlineData(SystemColor.Value.InactiveCaption, "123123")]
        [InlineData(SystemColor.Value.InactiveCaptionText, "321321")]
        [InlineData(SystemColor.Value.InfoBackground, "456456")]
        [InlineData(SystemColor.Value.InfoText, "654654")]
        [InlineData(SystemColor.Value.Menu, "789789")]
        [InlineData(SystemColor.Value.MenuBar, "987987")]
        [InlineData(SystemColor.Value.MenuHighlight, "147258")]
        [InlineData(SystemColor.Value.MenuText, "852963")]
        [InlineData(SystemColor.Value.ScrollBar, "369369")]
        [InlineData(SystemColor.Value.Window, "159159")]
        [InlineData(SystemColor.Value.WindowFrame, "951951")]
        [InlineData(SystemColor.Value.WindowText, "753753")]
        public void SystemColorReadWriteTest(SystemColor.Value colorValue, string lastColor)
        {
            Theme theme = new Theme("test");
            SystemColor color = new SystemColor(colorValue);
            color.LastColor = lastColor;
            theme.Colors.Dark1 = color;
            Workbook workbook = new Workbook();
            workbook.WorkbookTheme = theme;
            Assert.Equal(colorValue, ((SystemColor)workbook.WorkbookTheme.Colors.Dark1).ColorValue);
            Assert.Equal(lastColor, ((SystemColor)workbook.WorkbookTheme.Colors.Dark1).LastColor);
            Workbook givenWorkbook = TestUtils.WriteAndReadWorkbook(workbook);
            Assert.Equal(colorValue, ((SystemColor)givenWorkbook.WorkbookTheme.Colors.Dark1).ColorValue);
            Assert.Equal(lastColor, ((SystemColor)givenWorkbook.WorkbookTheme.Colors.Dark1).LastColor);
        }
    }
}
