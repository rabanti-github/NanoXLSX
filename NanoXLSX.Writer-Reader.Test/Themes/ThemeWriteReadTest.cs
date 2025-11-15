using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using NanoXLSX.Extensions;
using NanoXLSX.Interfaces;
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
            Theme theme = new Theme(name);
            Workbook workbook = new Workbook();
            workbook.WorkbookTheme = theme;
            Workbook givenWorkbook = TestUtils.WriteAndReadWorkbook(workbook);

            Assert.Equal(expectedName, givenWorkbook.WorkbookTheme.Name);
        }

        [Fact(DisplayName = "Test of the get and set function of the Colors property when saving and loading a workbook")]
        public void ColorsTest()
        {
            Theme theme = new Theme("test");
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

        [Fact(DisplayName = "Test of the get and set function Name property of the themes color scheme, when saving and loading a workbook")]
        public void ColorSchemeNameTest()
        {
            Theme theme = new Theme("test");
            Assert.NotNull(theme.Colors);
            Assert.Equal("default", theme.Colors.Name);
            theme.Colors.Name = "test1";
            Assert.Equal("test1", theme.Colors.Name);
            Workbook workbook = new Workbook();
            workbook.WorkbookTheme = theme;
            Workbook givenWorkbook = TestUtils.WriteAndReadWorkbook(workbook);
            Assert.Equal("test1", givenWorkbook.WorkbookTheme.Colors.Name);
        }


        [Theory(DisplayName = "Test of the get and set function of the themes color property 'Dark1' when saving and loading a workbook")]
        [MemberData(nameof(ColorTestData))]
        public void Dark1Test(object colorValue)
        {
            AssertColorProperty(colorValue,
                              (theme, color) => theme.Colors.Dark1 = color,
                              theme => theme.Colors.Dark1);
        }

        [Theory(DisplayName = "Test of the get and set function of the themes color property 'Light1' when saving and loading a workbook")]
        [MemberData(nameof(ColorTestData))]
        public void Light1Test(object colorValue)
        {
            AssertColorProperty(colorValue,
                              (theme, color) => theme.Colors.Light1 = color,
                              theme => theme.Colors.Light1);
        }

        [Theory(DisplayName = "Test of the get and set function of the themes color property 'Dark2' when saving and loading a workbook")]
        [MemberData(nameof(ColorTestData))]
        public void Dark2Test(object colorValue)
        {
            AssertColorProperty(colorValue,
                              (theme, color) => theme.Colors.Dark2 = color,
                              theme => theme.Colors.Dark2);
        }

        [Theory(DisplayName = "Test of the get and set function of the themes color property 'Light2' when saving and loading a workbook")]
        [MemberData(nameof(ColorTestData))]
        public void Light2Test(object colorValue)
        {
            AssertColorProperty(colorValue,
                              (theme, color) => theme.Colors.Light2 = color,
                              theme => theme.Colors.Light2);
        }

        [Theory(DisplayName = "Test of the get and set function of the themes color property 'Accent1' when saving and loading a workbook")]
        [MemberData(nameof(ColorTestData))]
        public void Accent1Test(object colorValue)
        {
            AssertColorProperty(colorValue,
                              (theme, color) => theme.Colors.Accent1 = color,
                              theme => theme.Colors.Accent1);
        }

        [Theory(DisplayName = "Test of the get and set function of the themes color property 'Accent2' when saving and loading a workbook")]
        [MemberData(nameof(ColorTestData))]
        public void Accent2Test(object colorValue)
        {
            AssertColorProperty(colorValue,
                              (theme, color) => theme.Colors.Accent2 = color,
                              theme => theme.Colors.Accent2);
        }

        [Theory(DisplayName = "Test of the get and set function of the themes color property 'Accent3' when saving and loading a workbook")]
        [MemberData(nameof(ColorTestData))]
        public void Accent3Test(object colorValue)
        {
            AssertColorProperty(colorValue,
                              (theme, color) => theme.Colors.Accent3 = color,
                              theme => theme.Colors.Accent3);
        }

        [Theory(DisplayName = "Test of the get and set function of the themes color property 'Accent4' when saving and loading a workbook")]
        [MemberData(nameof(ColorTestData))]
        public void Accent4Test(object colorValue)
        {
            AssertColorProperty(colorValue,
                              (theme, color) => theme.Colors.Accent4 = color,
                              theme => theme.Colors.Accent4);
        }

        [Theory(DisplayName = "Test of the get and set function of the themes color property 'Accent5' when saving and loading a workbook")]
        [MemberData(nameof(ColorTestData))]
        public void Accent5Test(object colorValue)
        {
            AssertColorProperty(colorValue,
                              (theme, color) => theme.Colors.Accent5 = color,
                              theme => theme.Colors.Accent5);
        }

        [Theory(DisplayName = "Test of the get and set function of the themes color property 'Accent6' when saving and loading a workbook")]
        [MemberData(nameof(ColorTestData))]
        public void Accent6Test(object colorValue)
        {
            AssertColorProperty(colorValue,
                              (theme, color) => theme.Colors.Accent6 = color,
                              theme => theme.Colors.Accent6);
        }

        [Theory(DisplayName = "Test of the get and set function of the themes color property 'Hyperlink' when saving and loading a workbook")]
        [MemberData(nameof(ColorTestData))]
        public void HyperlinkTest(object colorValue)
        {
            AssertColorProperty(colorValue,
                              (theme, color) => theme.Colors.Hyperlink = color,
                              theme => theme.Colors.Hyperlink);
        }

        [Theory(DisplayName = "Test of the get and set function of the themes color property 'FollowedHyperlink' when saving and loading a workbook")]
        [MemberData(nameof(ColorTestData))]
        public void FollowedHyperlinkTest(object colorValue)
        {
            AssertColorProperty(colorValue,
                              (theme, color) => theme.Colors.FollowedHyperlink = color,
                              theme => theme.Colors.FollowedHyperlink);
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


        // Test data for IColor properties.
        [ExcludeFromCodeCoverage]
        public static IEnumerable<object[]> ColorTestData =>
            new List<object[]>
            {
                new object[] { "FFFFFF" },
                new object[] { "000000" },
                new object[] { "ABCDEF" },
                new object[] { SystemColor.Value.ActiveCaption },
                new object[] { SystemColor.Value.GrayText },
                new object[] { SystemColor.Value.Window }
            };

        private void AssertColorProperty(object colorValue,
                                      System.Action<Theme, IColor> setColor,
                                      System.Func<Theme, IColor> getColor)
        {
            Theme theme = new Theme("test");
            IColor color;
            if (colorValue is string s)
            {
                color = new SrgbColor(s);
            }
            else
            {
                color = new SystemColor((SystemColor.Value)colorValue);
            }
            // Set the color using the passed lambda.
            setColor(theme, color);

            Workbook workbook = new Workbook();
            workbook.WorkbookTheme = theme;
            Workbook givenWorkbook = TestUtils.WriteAndReadWorkbook(workbook);

            // Assert that the saved and reloaded property matches.
            Assert.Equal(color, getColor(givenWorkbook.WorkbookTheme));
        }

    }
}
