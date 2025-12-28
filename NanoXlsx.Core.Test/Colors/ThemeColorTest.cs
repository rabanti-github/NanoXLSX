using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NanoXLSX.Colors;
using NanoXLSX.Themes;
using Xunit;

namespace NanoXLSX.Core.Test.Colors
{
    public class ThemeColorTest
    {
        [Fact(DisplayName = "Default Constructor Test")]
        public void ConstructorTest()
        {
            ThemeColor themeColor = new ThemeColor();
            Assert.Equal(Theme.ColorSchemeElement.Dark1, themeColor.ColorValue); // Implicit default value of enum
        }

        [Theory(DisplayName = "Default Constructor Test with arguments")]
        [InlineData(Theme.ColorSchemeElement.Dark1)]
        [InlineData(Theme.ColorSchemeElement.Light1)]
        [InlineData(Theme.ColorSchemeElement.Dark2)]
        [InlineData(Theme.ColorSchemeElement.Light2)]
        [InlineData(Theme.ColorSchemeElement.Accent1)]
        [InlineData(Theme.ColorSchemeElement.Accent2)]
        [InlineData(Theme.ColorSchemeElement.Accent3)]
        [InlineData(Theme.ColorSchemeElement.Accent4)]
        [InlineData(Theme.ColorSchemeElement.Accent5)]
        [InlineData(Theme.ColorSchemeElement.Accent6)]
        [InlineData(Theme.ColorSchemeElement.Hyperlink)]
        [InlineData(Theme.ColorSchemeElement.FollowedHyperlink)]
        public void ConstructorTest2(Theme.ColorSchemeElement colorSchemeElement)
        {
            ThemeColor themeColor = new ThemeColor(colorSchemeElement);
            Assert.Equal(colorSchemeElement, themeColor.ColorValue);
        }

        [Theory(DisplayName = "Test of the StringValue property")]
        [InlineData(Theme.ColorSchemeElement.Dark1, "0")]
        [InlineData(Theme.ColorSchemeElement.Light1, "1")]
        [InlineData(Theme.ColorSchemeElement.Dark2, "2")]
        [InlineData(Theme.ColorSchemeElement.Light2, "3")]
        [InlineData(Theme.ColorSchemeElement.Accent1, "4")]
        [InlineData(Theme.ColorSchemeElement.Accent2, "5")]
        [InlineData(Theme.ColorSchemeElement.Accent3, "6")]
        [InlineData(Theme.ColorSchemeElement.Accent4, "7")]
        [InlineData(Theme.ColorSchemeElement.Accent5, "8")]
        [InlineData(Theme.ColorSchemeElement.Accent6, "9")]
        [InlineData(Theme.ColorSchemeElement.Hyperlink, "10")]
        [InlineData(Theme.ColorSchemeElement.FollowedHyperlink, "11")]
        public void StringValueTest(Theme.ColorSchemeElement colorSchemeElement, string expectedValue)
        {
            ThemeColor themeColor = new ThemeColor(colorSchemeElement);
            Assert.Equal(expectedValue, themeColor.StringValue);
        }

        [Fact(DisplayName = "Test of the Equals method on equality (multiple cases)")]
        public void EqualsTestTrue()
        {
            ThemeColor color1 = new ThemeColor(Theme.ColorSchemeElement.Accent3);
            ThemeColor color2 = new ThemeColor(Theme.ColorSchemeElement.Accent3);
            Assert.True(color1.Equals(color2));

            ThemeColor color3 = new ThemeColor(Theme.ColorSchemeElement.Dark1);
            ThemeColor color4 = new ThemeColor(Theme.ColorSchemeElement.Dark1);
            Assert.True(color3.Equals(color4));
        }

        [Fact(DisplayName = "Test of the Equals method on inequality (multiple cases)")]
        public void EqualsTestFalse()
        {
            ThemeColor color1 = new ThemeColor(Theme.ColorSchemeElement.Accent3);
            ThemeColor color2 = new ThemeColor(Theme.ColorSchemeElement.Accent4);
            Assert.False(color1.Equals(color2));

            ThemeColor color3 = new ThemeColor(Theme.ColorSchemeElement.Dark1);
            ThemeColor color4 = new ThemeColor(Theme.ColorSchemeElement.Light1);
            Assert.False(color3.Equals(color4));
        }

        [Fact(DisplayName = "Test of the GetHashCode method on equality (multiple cases)")]
        public void GetHashCodeTestTrue()
        {
            ThemeColor color1 = new ThemeColor(Theme.ColorSchemeElement.Accent3);
            ThemeColor color2 = new ThemeColor(Theme.ColorSchemeElement.Accent3);
            Assert.Equal(color1.GetHashCode(), color2.GetHashCode());

            ThemeColor color3 = new ThemeColor(Theme.ColorSchemeElement.Dark1);
            ThemeColor color4 = new ThemeColor(Theme.ColorSchemeElement.Dark1);
            Assert.Equal(color3.GetHashCode(), color4.GetHashCode());
        }

        [Fact(DisplayName = "Test of the GetHashCode method on inequality (multiple cases)")]
        public void GetHashCodeTestFalse()
        {
            ThemeColor color1 = new ThemeColor(Theme.ColorSchemeElement.Accent3);
            ThemeColor color2 = new ThemeColor(Theme.ColorSchemeElement.Dark1);
            Assert.NotEqual(color1.GetHashCode(), color2.GetHashCode());

            ThemeColor color3 = new ThemeColor(Theme.ColorSchemeElement.Dark1);
            ThemeColor color4 = new ThemeColor(Theme.ColorSchemeElement.Light1);
            Assert.NotEqual(color3.GetHashCode(), color4.GetHashCode());
        }

    }
}
