using NanoXLSX.Interfaces;
using NanoXLSX.Themes;
using Xunit;

namespace NanoXLSX.Test.Core.ThemeTest
{
    public class ColorSchemeTest
    {


        [Theory(DisplayName = "Test of the get and set function of the Name property")]
        [InlineData("XYZ")]
        [InlineData(" ")]
        [InlineData("")]
        [InlineData(null)]
        public void NameTest(string value)
        {
            ColorScheme scheme = new ColorScheme();
            Assert.Null(scheme.Name);
            scheme.Name = value;
            Assert.Equal(value, scheme.Name);
        }

        [Theory(DisplayName = "Test of the get and set function of the color properties for SRGB values")]
        [InlineData("FFFFFF")]
        [InlineData("000000")]
        [InlineData("AACC3D")]
        public void ColorPropertiesSrgbTest(string value)
        {
            SrgbColor color = new SrgbColor(value);
            AssertColorProperties(color);
        }

        [Theory(DisplayName = "Test of the get and set function of the color properties for SystemColor values")]
        [InlineData(SystemColor.Value.ActiveBorder)]
        [InlineData(SystemColor.Value.ActiveCaption)]
        [InlineData(SystemColor.Value.AppWorkspace)]
        [InlineData(SystemColor.Value.Background)]
        [InlineData(SystemColor.Value.ButtonFace)]
        [InlineData(SystemColor.Value.ButtonHighlight)]
        [InlineData(SystemColor.Value.ButtonShadow)]
        [InlineData(SystemColor.Value.ButtonText)]
        [InlineData(SystemColor.Value.CaptionText)]
        [InlineData(SystemColor.Value.GradientActiveCaption)]
        [InlineData(SystemColor.Value.GradientInactiveCaption)]
        [InlineData(SystemColor.Value.GrayText)]
        [InlineData(SystemColor.Value.Highlight)]
        [InlineData(SystemColor.Value.HighlightText)]
        [InlineData(SystemColor.Value.HotLight)]
        [InlineData(SystemColor.Value.InactiveBorder)]
        [InlineData(SystemColor.Value.InactiveCaption)]
        [InlineData(SystemColor.Value.InactiveCaptionText)]
        [InlineData(SystemColor.Value.InfoBackground)]
        [InlineData(SystemColor.Value.InfoText)]
        [InlineData(SystemColor.Value.Menu)]
        [InlineData(SystemColor.Value.MenuBar)]
        [InlineData(SystemColor.Value.MenuHighlight)]
        [InlineData(SystemColor.Value.MenuText)]
        [InlineData(SystemColor.Value.ScrollBar)]
        [InlineData(SystemColor.Value.ThreeDimensionalDarkShadow)]
        [InlineData(SystemColor.Value.ThreeDimensionalLight)]
        [InlineData(SystemColor.Value.Window)]
        [InlineData(SystemColor.Value.WindowFrame)]
        [InlineData(SystemColor.Value.WindowText)]
        public void ColorPropertiesSystemColorTest(SystemColor.Value value)
        {
            SystemColor color = new SystemColor(value);
            AssertColorProperties(color);
        }

        [Fact(DisplayName = "Test of the get and set function of the color properties for null values")]
        public void PropertiesNullTest()
        {
            AssertColorProperties(null);
        }

        [Fact(DisplayName = "Test of Equals() and HashCode() implementations for equality")]
        public void EqualsTest()
        {
            ColorScheme scheme1 = new ColorScheme();
            scheme1.Name = "scheme1"; // Should have an influence
            scheme1.Dark1 = new SystemColor(SystemColor.Value.ActiveBorder);
            scheme1.Light1 = new SystemColor(SystemColor.Value.Menu);
            scheme1.Dark2 = new SystemColor(SystemColor.Value.Background);
            scheme1.Light2 = new SystemColor(SystemColor.Value.Background);
            scheme1.Accent1 = new SystemColor(SystemColor.Value.AppWorkspace);
            scheme1.Accent2 = new SystemColor(SystemColor.Value.ButtonShadow);
            scheme1.Accent3 = new SrgbColor("FFAABB");
            scheme1.Accent4 = null;
            scheme1.Accent5 = new SrgbColor("FFAABB");
            scheme1.Accent6 = new SrgbColor("FFAABB");
            scheme1.Hyperlink = new SrgbColor("FFAABB");
            scheme1.FollowedHyperlink = new SrgbColor("FFAABB");

            ColorScheme scheme2 = new ColorScheme();
            scheme2.Name = "scheme1"; // Should have an influence
            scheme2.Dark1 = new SystemColor(SystemColor.Value.ActiveBorder);
            scheme2.Light1 = new SystemColor(SystemColor.Value.Menu);
            scheme2.Dark2 = new SystemColor(SystemColor.Value.Background);
            scheme2.Light2 = new SystemColor(SystemColor.Value.Background);
            scheme2.Accent1 = new SystemColor(SystemColor.Value.AppWorkspace);
            scheme2.Accent2 = new SystemColor(SystemColor.Value.ButtonShadow);
            scheme2.Accent3 = new SrgbColor("FFAABB");
            scheme2.Accent4 = null;
            scheme2.Accent5 = new SrgbColor("FFAABB");
            scheme2.Accent6 = new SrgbColor("FFAABB");
            scheme2.Hyperlink = new SrgbColor("FFAABB");
            scheme2.FollowedHyperlink = new SrgbColor("FFAABB");

            Assert.True(scheme1.Equals(scheme2));
            Assert.Equal(scheme1.GetHashCode(), scheme2.GetHashCode());
        }

        [Fact(DisplayName = "Test Equals method for Theme")]
        public void ThemeEqualsTest()
        {
            var theme1 = new Theme("TestTheme");
            var theme2 = new Theme("TestTheme");

            IColor newDark1 = new SrgbColor("123456");
            IColor newAccent1 = new SrgbColor("654321");
            IColor newHyperlink = new SrgbColor("ABC123");

            theme1.Colors.Dark1 = newDark1;
            theme1.Colors.Accent1 = newAccent1;
            theme1.Colors.Hyperlink = newHyperlink;

            theme2.Colors.Dark1 = newDark1;
            theme2.Colors.Accent1 = newAccent1;
            theme2.Colors.Hyperlink = newHyperlink;

            Assert.True(theme1.Equals(theme2));
        }

        [Fact(DisplayName = "Test GetHashCode method for Theme")]
        public void ThemeGetHashCodeTest()
        {
            var theme1 = new Theme("TestTheme");
            var theme2 = new Theme("TestTheme");

            IColor newDark1 = new SrgbColor("123456");
            IColor newAccent1 = new SrgbColor("654321");
            IColor newHyperlink = new SrgbColor("ABC123");

            theme1.Colors.Dark1 = newDark1;
            theme1.Colors.Accent1 = newAccent1;
            theme1.Colors.Hyperlink = newHyperlink;

            theme2.Colors.Dark1 = newDark1;
            theme2.Colors.Accent1 = newAccent1;
            theme2.Colors.Hyperlink = newHyperlink;

            Assert.Equal(theme1.GetHashCode(), theme2.GetHashCode());
        }

        private void AssertColorProperties(IColor color)
        {
            ColorScheme scheme = new ColorScheme();
            Assert.Null(scheme.Dark1);
            Assert.Null(scheme.Light1);
            Assert.Null(scheme.Dark2);
            Assert.Null(scheme.Light2);
            Assert.Null(scheme.Accent1);
            Assert.Null(scheme.Accent2);
            Assert.Null(scheme.Accent3);
            Assert.Null(scheme.Accent4);
            Assert.Null(scheme.Accent5);
            Assert.Null(scheme.Accent6);
            Assert.Null(scheme.Hyperlink);
            Assert.Null(scheme.FollowedHyperlink);
            scheme.Dark1 = color;
            scheme.Light1 = color;
            scheme.Dark2 = color;
            scheme.Light2 = color;
            scheme.Accent1 = color;
            scheme.Accent2 = color;
            scheme.Accent3 = color;
            scheme.Accent4 = color;
            scheme.Accent5 = color;
            scheme.Accent6 = color;
            scheme.Hyperlink = color;
            scheme.FollowedHyperlink = color;
            Assert.Equal(color, scheme.Dark1);
            Assert.Equal(color, scheme.Light1);
            Assert.Equal(color, scheme.Dark2);
            Assert.Equal(color, scheme.Light2);
            Assert.Equal(color, scheme.Accent1);
            Assert.Equal(color, scheme.Accent2);
            Assert.Equal(color, scheme.Accent3);
            Assert.Equal(color, scheme.Accent4);
            Assert.Equal(color, scheme.Accent5);
            Assert.Equal(color, scheme.Accent6);
            Assert.Equal(color, scheme.Hyperlink);
            Assert.Equal(color, scheme.FollowedHyperlink);
        }
    }
}
