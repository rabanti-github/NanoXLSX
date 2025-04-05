using System.Collections.Generic;
using NanoXLSX.Exceptions;
using NanoXLSX.Interfaces;
using NanoXLSX.Themes;
using Xunit;

namespace NanoXLSX.Core.Test.Themes
{
    public class SystemColorTest
    {

        [Theory(DisplayName = "Test of the getter and setter of the ColorValue property")]
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
        public void ColorValueTest(SystemColor.Value value)
        {
            SystemColor color = new SystemColor();
            Assert.Equal(SystemColor.Value.WindowText, color.ColorValue); // Default
            color.ColorValue = value;
            Assert.Equal(value, color.ColorValue);
        }

        [Theory(DisplayName = "Test of the getter of the StringValue property")]
        [InlineData(SystemColor.Value.ActiveBorder, "activeBorder")]
        [InlineData(SystemColor.Value.ActiveCaption, "activeCaption")]
        [InlineData(SystemColor.Value.AppWorkspace, "appWorkspace")]
        [InlineData(SystemColor.Value.Background, "background")]
        [InlineData(SystemColor.Value.ButtonFace, "btnFace")]
        [InlineData(SystemColor.Value.ButtonHighlight, "btnHighlight")]
        [InlineData(SystemColor.Value.ButtonShadow, "btnShadow")]
        [InlineData(SystemColor.Value.ButtonText, "btnText")]
        [InlineData(SystemColor.Value.CaptionText, "captionText")]
        [InlineData(SystemColor.Value.GradientActiveCaption, "gradientActiveCaption")]
        [InlineData(SystemColor.Value.GradientInactiveCaption, "gradientInactiveCaption")]
        [InlineData(SystemColor.Value.GrayText, "grayText")]
        [InlineData(SystemColor.Value.Highlight, "highlight")]
        [InlineData(SystemColor.Value.HighlightText, "highlightText")]
        [InlineData(SystemColor.Value.HotLight, "hotLight")]
        [InlineData(SystemColor.Value.InactiveBorder, "inactiveBorder")]
        [InlineData(SystemColor.Value.InactiveCaption, "inactiveCaption")]
        [InlineData(SystemColor.Value.InactiveCaptionText, "inactiveCaptionText")]
        [InlineData(SystemColor.Value.InfoBackground, "infoBk")]
        [InlineData(SystemColor.Value.InfoText, "infoText")]
        [InlineData(SystemColor.Value.Menu, "menu")]
        [InlineData(SystemColor.Value.MenuBar, "menuBar")]
        [InlineData(SystemColor.Value.MenuHighlight, "menuHighlight")]
        [InlineData(SystemColor.Value.MenuText, "menuText")]
        [InlineData(SystemColor.Value.ScrollBar, "scrollBar")]
        [InlineData(SystemColor.Value.ThreeDimensionalDarkShadow, "3dDkShadow")]
        [InlineData(SystemColor.Value.ThreeDimensionalLight, "3dLight")]
        [InlineData(SystemColor.Value.Window, "window")]
        [InlineData(SystemColor.Value.WindowFrame, "windowFrame")]
        [InlineData(SystemColor.Value.WindowText, "windowText")]
        public void StringValueTest(SystemColor.Value givenValue, string expectedValue)
        {
            SystemColor color = new SystemColor();
            Assert.Equal(SystemColor.Value.WindowText, color.ColorValue); // Default
            color.ColorValue = givenValue;
            Assert.Equal(expectedValue, color.StringValue);
        }

        [Fact(DisplayName = "Test of the failing StringValue property on invalid values")]
        public void StringValueFailTest()
        {
            SystemColor color = new SystemColor((SystemColor.Value)99);
            Assert.Throws<StyleException>(() => color.StringValue);
        }

        [Theory(DisplayName = "Test of the getter and setter of the LastColor property on valid values")]
        [InlineData("FFFFFF")]
        [InlineData("000000")]
        [InlineData("ABCDEF")]
        [InlineData("123456")]
        [InlineData("abcdef")]
        [InlineData("ffaabb")]
        public void LastColorTest(string srgbValue)
        {
            SystemColor color = new SystemColor();
            Assert.Equal("000000", color.LastColor); // Default black
            color.LastColor = srgbValue;
            Assert.Equal(srgbValue, color.LastColor);
        }

        [Theory(DisplayName = "Test of the failing getter and setter of the LastColor property on invalid values")]
        [InlineData("-1")]
        [InlineData("0")]
        [InlineData("")]
        [InlineData(null)]
        [InlineData("XABBCC")]
        [InlineData("AAAAA")]
        [InlineData("AAAAAAA")]
        [InlineData("AAAAAAAA")]
        [InlineData("01234")]
        [InlineData("#001122")]
        [InlineData("-aabbcc")]
        public void LastColorFailTest(string srgbValue)
        {
            SystemColor color = new SystemColor();
            Assert.Equal("000000", color.LastColor); // Default black
            Assert.Throws<StyleException>(() => color.LastColor = srgbValue);
        }

        [Theory(DisplayName = "Test of the constructor with the color value as argument")]
        [InlineData(SystemColor.Value.ActiveBorder, "AABBCC")]
        [InlineData(SystemColor.Value.ActiveCaption, "FFFFFF")]
        [InlineData(SystemColor.Value.AppWorkspace, "000000")]
        [InlineData(SystemColor.Value.Background, "999999")]
        [InlineData(SystemColor.Value.ButtonFace, "A3F4C5")]
        [InlineData(SystemColor.Value.ButtonHighlight, "aaaaaa")]
        [InlineData(SystemColor.Value.ButtonShadow, "ffffff")]
        [InlineData(SystemColor.Value.ButtonText, "012345")]
        [InlineData(SystemColor.Value.CaptionText, "A9A9A9")]
        [InlineData(SystemColor.Value.GradientActiveCaption, "A1c4F9")]
        [InlineData(SystemColor.Value.GradientInactiveCaption, "000001")]
        [InlineData(SystemColor.Value.GrayText, "100000")]
        [InlineData(SystemColor.Value.Highlight, "ABCDEF")]
        [InlineData(SystemColor.Value.HighlightText, "aabbcc")]
        [InlineData(SystemColor.Value.HotLight, "ffffff")]
        [InlineData(SystemColor.Value.InactiveBorder, "010101")]
        [InlineData(SystemColor.Value.InactiveCaption, "a4a4a4")]
        [InlineData(SystemColor.Value.InactiveCaptionText, "CCCCCC")]
        [InlineData(SystemColor.Value.InfoBackground, "BbBbBb")]
        [InlineData(SystemColor.Value.InfoText, "898900")]
        [InlineData(SystemColor.Value.Menu, "cccccc")]
        [InlineData(SystemColor.Value.MenuBar, "0A0B0C")]
        [InlineData(SystemColor.Value.MenuHighlight, "777777")]
        [InlineData(SystemColor.Value.MenuText, "70A9f7")]
        [InlineData(SystemColor.Value.ScrollBar, "4cff33")]
        [InlineData(SystemColor.Value.ThreeDimensionalDarkShadow, "00000A")]
        [InlineData(SystemColor.Value.ThreeDimensionalLight, "FFFFFE")]
        [InlineData(SystemColor.Value.Window, "eeeeef")]
        [InlineData(SystemColor.Value.WindowFrame, "65CC78")]
        [InlineData(SystemColor.Value.WindowText, "AD44FF")]
        public void ConstructorTest2(SystemColor.Value value, string lastColor)
        {
            SystemColor color = new SystemColor(value, lastColor);
            Assert.Equal(value, color.ColorValue);
            Assert.Equal(lastColor, color.LastColor);
        }

        [Theory(DisplayName = "Test of the failing constructor with arguments (color value and last color) on invalid values")]
        [InlineData(SystemColor.Value.InactiveCaptionText, "-1")]
        [InlineData(SystemColor.Value.WindowText, "0")]
        [InlineData(SystemColor.Value.GrayText, "")]
        [InlineData(SystemColor.Value.ActiveBorder, null)]
        [InlineData(SystemColor.Value.HighlightText, "XABBCC")]
        [InlineData(SystemColor.Value.Background, "AAAAA")]
        [InlineData(SystemColor.Value.ButtonShadow, "AAAAAAA")]
        [InlineData(SystemColor.Value.CaptionText, "AAAAAAAA")]
        [InlineData(SystemColor.Value.ButtonHighlight, "01234")]
        [InlineData(SystemColor.Value.ActiveCaption, "#001122")]
        [InlineData(SystemColor.Value.ButtonFace, "-aabbcc")]
        public void ConstructorFailTest(SystemColor.Value value, string srgbValue)
        {
            Assert.Throws<StyleException>(() => new SystemColor(value, srgbValue));
        }

        [Fact(DisplayName = "Test of the Equals method (multiple cases)")]
        public void EqualsTest()
        {
            SystemColor color1 = new SystemColor(SystemColor.Value.ButtonHighlight);
            color1.LastColor = "112233";
            SystemColor color2 = new SystemColor();
            color2.ColorValue = SystemColor.Value.ButtonHighlight;
            color2.LastColor = "112233";
            Assert.True(color1.Equals(color2));

            SystemColor color3 = new SystemColor();
            SystemColor color4 = new SystemColor();
            Assert.True(color3.Equals(color4));
        }

        [Fact(DisplayName = "Test of the Equals method on inequality (multiple cases)")]
        public void EqualsTest2()
        {
            SystemColor color1 = new SystemColor(SystemColor.Value.CaptionText);
            SystemColor color2 = new SystemColor();
            color2.ColorValue = SystemColor.Value.GradientActiveCaption;
            Assert.False(color1.Equals(color2));

            SystemColor color3 = new SystemColor(SystemColor.Value.ActiveCaption);
            SystemColor color4 = new SystemColor();
            Assert.False(color3.Equals(color4));

            SystemColor color5 = new SystemColor();
            DummyColor color6 = new DummyColor();
            Assert.False(color5.Equals(color6));

            SystemColor color7 = new SystemColor(SystemColor.Value.CaptionText, "AABBCC");
            SystemColor color8 = new SystemColor(SystemColor.Value.CaptionText, "001122");
            Assert.False(color7.Equals(color8));
        }

        [Fact(DisplayName = "Test of the GetHashCode method (multiple cases)")]
        public void GetHashCodeTest()
        {
            SystemColor color1 = new SystemColor(SystemColor.Value.AppWorkspace);
            SystemColor color2 = new SystemColor();
            color2.ColorValue = SystemColor.Value.AppWorkspace;
            Assert.Equal(color1.GetHashCode(), color2.GetHashCode());

            SystemColor color3 = new SystemColor();
            SystemColor color4 = new SystemColor();
            Assert.Equal(color3.GetHashCode(), color4.GetHashCode());

            SystemColor color5 = new SystemColor(SystemColor.Value.AppWorkspace, "CCDDEE");
            SystemColor color6 = new SystemColor();
            color6.ColorValue = SystemColor.Value.AppWorkspace;
            color6.LastColor = "CCDDEE";
            Assert.Equal(color5.GetHashCode(), color6.GetHashCode());
        }

        [Fact(DisplayName = "Test of the GetHashCode method on inequality (multiple cases)")]
        public void GetHashCodeTest2()
        {
            SystemColor color1 = new SystemColor(SystemColor.Value.Background);
            SystemColor color2 = new SystemColor();
            color2.ColorValue = SystemColor.Value.ButtonFace;
            Assert.NotEqual(color1.GetHashCode(), color2.GetHashCode());

            SystemColor color3 = new SystemColor(SystemColor.Value.AppWorkspace);
            SystemColor color4 = new SystemColor();
            Assert.NotEqual(color3.GetHashCode(), color4.GetHashCode());

            SystemColor color5 = new SystemColor();
            DummyColor color6 = new DummyColor();
            Assert.NotEqual(color5.GetHashCode(), color6.GetHashCode());

            SystemColor color7 = new SystemColor(SystemColor.Value.Background, "AACCDD");
            SystemColor color8 = new SystemColor();
            color8.ColorValue = SystemColor.Value.Background;
            color8.LastColor = "002233";
            Assert.NotEqual(color7.GetHashCode(), color8.GetHashCode());
        }

        private class DummyColor : IColor
        {
            public string StringValue => null;

            public override int GetHashCode()
            {
                return 800285906 + EqualityComparer<string>.Default.GetHashCode(StringValue);
            }
        }


    }
}
