using System.Collections.Generic;
using NanoXLSX.Colors;
using NanoXLSX.Exceptions;
using NanoXLSX.Interfaces;
using Xunit;

namespace NanoXLSX.Core.Test.Colors
{
    public class SrgbColorTest
    {

        [Theory(DisplayName = "Test of the getter and setter of the ColorValue property on valid values")]
        [InlineData("FFFFFF", "FFFFFFFF")]
        [InlineData("000000", "FF000000")]
        [InlineData("ABCDEF", "FFABCDEF")]
        [InlineData("123456", "FF123456")]
        [InlineData("abcdef", "FFABCDEF")]
        [InlineData("ffaabb", "FFFFAABB")]
        public void ColorValueTest(string givenSrgbValue, string expectedSrgbValue)
        {
            var color = new SrgbColor();
            Assert.Null(color.ColorValue);
            color.ColorValue = givenSrgbValue;
            Assert.Equal(expectedSrgbValue, color.ColorValue);
        }

        [Theory(DisplayName = "Test of the failing getter and setter of the ColorValue property on invalid values")]
        [InlineData("-1")]
        [InlineData("0")]
        [InlineData("")]
        [InlineData(null)]
        [InlineData("XABBCC")]
        [InlineData("AAAAA")]
        [InlineData("AAAAAAA")]
        [InlineData("AAAAAAAAA")]
        [InlineData("#AAAAAAAA")]
        [InlineData("01234")]
        [InlineData("#001122")]
        [InlineData("-aabbcc")]
        public void ColorValueFailTest(string srgbValue)
        {
            var color = new SrgbColor();
            Assert.Null(color.ColorValue);
            Assert.Throws<StyleException>(() => color.ColorValue = srgbValue);
        }


        [Theory(DisplayName = "Test of the getter of the StringValue property on valid values")]
        [InlineData("FFFFFF", "FFFFFFFF")]
        [InlineData("000000", "FF000000")]
        [InlineData("ABCDEF", "FFABCDEF")]
        [InlineData("123456", "FF123456")]
        [InlineData("abcdef", "FFABCDEF")]
        [InlineData("ffaabb", "FFFFAABB")]
        public void StringValueTest(string givenSrgbValue, string expectedSrgbValue)
        {
            var color = new SrgbColor();
            Assert.Null(color.StringValue);
            color.ColorValue = givenSrgbValue;
            Assert.Equal(expectedSrgbValue, color.StringValue);
        }


        [Theory(DisplayName = "Test of Constructor with arguments (ColorValue) on valid values")]
        [InlineData("FFFFFF", "FFFFFFFF")]
        [InlineData("000000", "FF000000")]
        [InlineData("ABCDEF", "FFABCDEF")]
        [InlineData("123456", "FF123456")]
        [InlineData("abcdef", "FFABCDEF")]
        [InlineData("ffaabb", "FFFFAABB")]
        public void ConstructorTest(string givenSrgbValue, string expectedSrgbValue)
        {
            var color = new SrgbColor(givenSrgbValue);
            Assert.Equal(expectedSrgbValue, color.ColorValue);
        }

        [Theory(DisplayName = "Test of the failing constructor with arguments (ColorValue) on invalid values")]
        [InlineData("-1")]
        [InlineData("0")]
        [InlineData("")]
        [InlineData(null)]
        [InlineData("XABBCC")]
        [InlineData("AAAAA")]
        [InlineData("AAAAAAA")]
        [InlineData("AAAAAAAAA")]
        [InlineData("#AAAAAAAA")]
        [InlineData("01234")]
        [InlineData("#001122")]
        [InlineData("-aabbcc")]
        public void ConstructorFailTest(string srgbValue)
        {
            Assert.Throws<StyleException>(() => new SrgbColor(srgbValue));
        }

        [Theory(DisplayName = "Test of the ToArgbColor function")]
        [InlineData("FFFFFF", "FFFFFFFF")]
        [InlineData("000000", "FF000000")]
        [InlineData("ABCDEF", "FFABCDEF")]
        [InlineData("123456", "FF123456")]
        [InlineData("abcdef", "FFABCDEF")]
        [InlineData("ffaabb", "FFFFAABB")]
        public void ToArgbColorTest(string srgbValue, string expectedArgbColor)
        {
            var color = new SrgbColor(srgbValue);
            Assert.Equal(expectedArgbColor, color.ColorValue);
        }

        [Fact(DisplayName = "Test of the Equals method (multiple cases)")]
        public void EqualsTest()
        {
            var color1 = new SrgbColor("ACADAF");
            var color2 = new SrgbColor
            {
                ColorValue = "ACADAF"
            };
            Assert.True(color1.Equals(color2));

            var color3 = new SrgbColor();
            var color4 = new SrgbColor();
            Assert.True(color3.Equals(color4));
        }

        [Fact(DisplayName = "Test of the Equals method on inequality (multiple cases)")]
        public void EqualsTest2()
        {
            var color1 = new SrgbColor("ACADAF");
            var color2 = new SrgbColor
            {
                ColorValue = "ACADA0"
            };
            Assert.False(color1.Equals(color2));

            var color3 = new SrgbColor("ACADAF");
            var color4 = new SrgbColor();
            Assert.False(color3.Equals(color4));

            var color5 = new SrgbColor();
            var color6 = new DummyColor();
            Assert.False(color5.Equals(color6));
        }

        [Fact(DisplayName = "Test of the GetHashCode method (multiple cases)")]
        public void GetHashCodeTest()
        {
            var color1 = new SrgbColor("ACADAF");
            var color2 = new SrgbColor
            {
                ColorValue = "ACADAF"
            };
            Assert.Equal(color1.GetHashCode(), color2.GetHashCode());

            var color3 = new SrgbColor();
            var color4 = new SrgbColor();
            Assert.Equal(color3.GetHashCode(), color4.GetHashCode());
        }

        [Fact(DisplayName = "Test of the GetHashCode method on inequality (multiple cases)")]
        public void GetHashCodeTest2()
        {
            var color1 = new SrgbColor("ACADAF");
            var color2 = new SrgbColor
            {
                ColorValue = "ACADA0"
            };
            Assert.NotEqual(color1.GetHashCode(), color2.GetHashCode());

            var color3 = new SrgbColor("ACADAF");
            var color4 = new SrgbColor();
            Assert.NotEqual(color3.GetHashCode(), color4.GetHashCode());

            var color5 = new SrgbColor();
            var color6 = new DummyColor();
            Assert.NotEqual(color5.GetHashCode(), color6.GetHashCode());
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
