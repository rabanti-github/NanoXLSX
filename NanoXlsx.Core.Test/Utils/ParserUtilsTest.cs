using System;
using NanoXLSX.Utils;
using Xunit;

namespace NanoXLSX.Test.Core.UtilsTest
{
    public class ParserUtilsTest
    {
        [Theory(DisplayName = "Test of the ParserUtils ToUpper function")]
        [InlineData("", "")]
        [InlineData(null, null)]
        [InlineData("123", "123")]
        [InlineData("abc", "ABC")]
        [InlineData("ABC", "ABC")]
        public void ToUpperTest(string givenValue, string expectedValue)
        {
            string value = ParserUtils.ToUpper(givenValue);
            Assert.Equal(expectedValue, value);
        }

        [Theory(DisplayName = "Test of the ParserUtils StartsWith function")]
        [InlineData("HelloWorld", "Hello", true)]
        [InlineData("HelloWorld", "world", false)]
        [InlineData("000", "0", true)]
        [InlineData("", "", true)]
        [InlineData(null, null, true)]
        [InlineData(null, "test", false)]
        [InlineData("test", null, false)]
        [InlineData("012", "3", false)]
        [InlineData("abc", "abc", true)]
        [InlineData("abc", "ABC", false)]
        [InlineData("   ", " ", true)]
        [InlineData("   ", "\t", false)]
        public void StartsWithTest(string givenValue, string startValue, bool expectedStartsWith)
        {
            bool startsWith = ParserUtils.StartsWith(givenValue, startValue);
            Assert.Equal(expectedStartsWith, startsWith);
        }

        [Theory(DisplayName = "Test of the ParserUtils NotStartsWith function")]
        [InlineData("HelloWorld", "Hello", false)]
        [InlineData("HelloWorld", "world", true)]
        [InlineData("000", "0", false)]
        [InlineData("", "", false)]
        [InlineData(null, null, false)]
        [InlineData(null, "test", true)]
        [InlineData("test", null, true)]
        [InlineData("012", "3", true)]
        [InlineData("abc", "abc", false)]
        [InlineData("abc", "ABC", true)]
        [InlineData("   ", " ", false)]
        [InlineData("   ", "\t", true)]
        public void NotStartsWithTest(string givenValue, string startValue, bool expectedStartsWith)
        {
            bool startsWith = ParserUtils.NotStartsWith(givenValue, startValue);
            Assert.Equal(expectedStartsWith, startsWith);
        }

        [Theory(DisplayName = "Test of the ParserUtils ToString function for integers")]
        [InlineData(-10, "-10")]
        [InlineData(0, "0")]
        [InlineData(1, "1")]
        [InlineData(100, "100")]
        public void ToStringTest(int givenValue, string expectedValue)
        {
            string value = ParserUtils.ToString(givenValue);
            Assert.Equal(expectedValue, value);
        }

        [Theory(DisplayName = "Test of the ParserUtils ToString function for floats")]
        [InlineData(-10f, "-10")]
        [InlineData(0f, "0")]
        [InlineData(1f, "1")]
        [InlineData(100f, "100")]
        [InlineData(0.1f, "0.1")]
        [InlineData(-0.01f, "-0.01")]
        [InlineData(100.01f, "100.01")]
        [InlineData(-1.111f, "-1.111")]
        public void ToStringTest2(float givenValue, string expectedValue)
        {
            string value = ParserUtils.ToString(givenValue);
            Assert.Equal(expectedValue, value);
        }

        [Theory(DisplayName = "Test of the ParserUtils NormalizeNewLines function")]
        [InlineData(null, null)]
        [InlineData("", "")]
        [InlineData("test", "test")]
        [InlineData("test\r\ntest", "test\r\ntest")]
        [InlineData("test\rtest", "test\r\ntest")]
        [InlineData("test\ntest", "test\r\ntest")]
        [InlineData("test\n\rtest", "test\r\ntest")]
        [InlineData("test\r\ntest \r\ntest", "test\r\ntest \r\ntest")]
        [InlineData("test\rtest \rtest", "test\r\ntest \r\ntest")]
        [InlineData("test\ntest \ntest", "test\r\ntest \r\ntest")]
        [InlineData("test\n\rtest \n\rtest", "test\r\ntest \r\ntest")]
        [InlineData("\n\r\n\n", "\r\n\r\n\r\n")]
        public void NormalizeNewLinesTest(string givenValue, string expectedValue)
        {
            string value = ParserUtils.NormalizeNewLines(givenValue);
            Assert.Equal(expectedValue, value);
        }


        [Theory(DisplayName = "Test of the ParserUtils ParseFloat function (no error handling)")]
        [InlineData("1", 1f)]
        [InlineData("0", 0f)]
        [InlineData("-1", -1f)]
        [InlineData("-10", -10f)]
        [InlineData("22", 22f)]
        [InlineData("-0.005", -0.005)]
        [InlineData("0.858", 0.858f)]
        [InlineData("-99998.1234", -99998.1234f)]
        [InlineData("98755142.237", 98755142.237f)]
        public void ParseFloatTest(String givenValue, float expectedValue)
        {
            float value = ParserUtils.ParseFloat(givenValue);
            Assert.Equal(expectedValue, value);
        }

        [Theory(DisplayName = "Test of the ParserUtils ParseInt function (no error handling)")]
        [InlineData("0", 0)]
        [InlineData("1", 1)]
        [InlineData("1.0", 1)]
        [InlineData("-2.0", -2)]
        [InlineData("0.0", 0)]
        [InlineData("-1", -1)]
        [InlineData("42", 42)]
        [InlineData("-42", -42)]
        [InlineData("2147483647", int.MaxValue)]
        [InlineData("-2147483648", int.MinValue)]
        public void ParseIntTest(String givenValue, int expectedValue)
        {
            int value = ParserUtils.ParseInt(givenValue);
            Assert.Equal(expectedValue, value);
        }

        [Theory(DisplayName = "Test of the ParserUtils ParseBinaryBool function (no error handling)")]
        [InlineData("0", 0)]
        [InlineData("1", 1)]
        [InlineData("-1", 0)]
        [InlineData("2", 1)]
        [InlineData("false", 0)]
        [InlineData("FALSE", 0)]
        [InlineData("False", 0)]
        [InlineData("", 0)]
        [InlineData(null, 0)]
        [InlineData("no", 0)]
        [InlineData("true", 1)]
        [InlineData("TRUE", 1)]
        [InlineData("True", 1)]
        public void ParseBinaryBoolTest(String givenValue, int expectedValue)
        {
            int value = ParserUtils.ParseBinaryBool(givenValue);
            Assert.Equal(expectedValue, value);
        }

        [Theory(DisplayName = "Test of the ParserUtils TryParseFloat function")]
        [InlineData("1", 1f, true)]
        [InlineData("0", 0f, true)]
        [InlineData("-1", -1f, true)]
        [InlineData("-10", -10f, true)]
        [InlineData("22", 22f, true)]
        [InlineData("-0.005", -0.005f, true)]
        [InlineData("0.858", 0.858f, true)]
        [InlineData("-99998.1234", -99998.1234f, true)]
        [InlineData("98755142.237", 98755142.237f, true)]
        [InlineData("", 0f, false)]
        [InlineData(" ", 0f, false)]
        [InlineData(null, 0f, false)]
        [InlineData("a", 0f, false)]
        [InlineData("1x1", 0f, false)]
        [InlineData("0.0x", 0f, false)]
        [InlineData("-22.5f4", 0f, false)]
        public void TryParseFloatTest(String givenValue, float expectedValue, bool expectedMatch)
        {
            bool match = ParserUtils.TryParseFloat(givenValue, out var value);
            Assert.Equal(expectedValue, value);
            Assert.Equal(expectedMatch, match);
        }

        [Theory(DisplayName = "Test of the ParserUtils TryParseInt function")]
        [InlineData("0", 0, true)]
        [InlineData("1", 1, true)]
        [InlineData("-1", -1, true)]
        [InlineData("42", 42, true)]
        [InlineData("-42", -42, true)]
        [InlineData("2147483647", int.MaxValue, true)]
        [InlineData("", 0, false)]
        [InlineData(" ", 0, false)]
        [InlineData(null, 0, false)]
        [InlineData("a", 0, false)]
        [InlineData("1x1", 0, false)]
        public void TryParseIntTest(String givenValue, int expectedValue, bool expectedMatch)
        {
            bool match = ParserUtils.TryParseInt(givenValue, out var value);
            Assert.Equal(expectedValue, value);
            Assert.Equal(expectedMatch, match);
        }

        [Theory(DisplayName = "Test of the ParserUtils TryParseUint function")]
        [InlineData("0", 0, true)]
        [InlineData("1", 1, true)]
        [InlineData("42", 42, true)]
        [InlineData("2147483647", int.MaxValue, true)]
        [InlineData("4294967295", uint.MaxValue, true)]
        [InlineData("", 0, false)]
        [InlineData(" ", 0, false)]
        [InlineData(null, 0, false)]
        [InlineData("a", 0, false)]
        [InlineData("1x1", 0, false)]
        [InlineData("-1", 0, false)]
        [InlineData("-42", 0, false)]
        public void TryParseUintTest(String givenValue, uint expectedValue, bool expectedMatch)
        {
            bool match = ParserUtils.TryParseUint(givenValue, out var value);
            Assert.Equal(expectedValue, value);
            Assert.Equal(expectedMatch, match);
        }

        [Theory(DisplayName = "Test of the ParserUtils TryParseLong function")]
        [InlineData("0", 0, true)]
        [InlineData("1", 1, true)]
        [InlineData("-1", -1, true)]
        [InlineData("42", 42, true)]
        [InlineData("-42", -42, true)]
        [InlineData("9223372036854775807", long.MaxValue, true)]
        [InlineData("", 0, false)]
        [InlineData(" ", 0, false)]
        [InlineData(null, 0, false)]
        [InlineData("a", 0, false)]
        [InlineData("1x1", 0, false)]
        public void TryParseLongTest(String givenValue, long expectedValue, bool expectedMatch)
        {
            bool match = ParserUtils.TryParseLong(givenValue, out var value);
            Assert.Equal(expectedValue, value);
            Assert.Equal(expectedMatch, match);
        }

        [Theory(DisplayName = "Test of the ParserUtils TryParseUlong function")]
        [InlineData("0", 0, true)]
        [InlineData("1", 1, true)]
        [InlineData("42", 42, true)]
        [InlineData("9223372036854775807", long.MaxValue, true)]
        [InlineData("18446744073709551615", ulong.MaxValue, true)]
        [InlineData("", 0, false)]
        [InlineData(" ", 0, false)]
        [InlineData(null, 0, false)]
        [InlineData("a", 0, false)]
        [InlineData("1x1", 0, false)]
        [InlineData("-1", 0, false)]
        [InlineData("-42", 0, false)]
        public void TryParseUlongTest(string givenValue, ulong expectedValue, bool expectedMatch)
        {
            bool match = ParserUtils.TryParseUlong(givenValue, out var value);
            Assert.Equal(expectedValue, value);
            Assert.Equal(expectedMatch, match);
        }

        [Theory(DisplayName = "Test of the ParserUtils TryParseDecimal function")]
        [InlineData("1", 1, true)]
        [InlineData("0", 0, true)]
        [InlineData("-1", -1, true)]
        [InlineData("-10", -10, true)]
        [InlineData("22", 22, true)]
        [InlineData("-0.0000005", -0.0000005, true)]
        [InlineData("0.858", 0.858, true)]
        [InlineData("-99998.1234", -99998.1234, true)]
        [InlineData("98755142.2111137", 98755142.2111137, true)]
        [InlineData("", 0, false)]
        [InlineData(" ", 0, false)]
        [InlineData(null, 0, false)]
        [InlineData("a", 0, false)]
        [InlineData("1x1", 0, false)]
        [InlineData("0.0x", 0, false)]
        [InlineData("-22.5f4", 0, false)]
        public void TryParseDecimalTest(string givenValue, decimal expectedValue, bool expectedMatch)
        {
            bool match = ParserUtils.TryParseDecimal(givenValue, out var value);
            Assert.Equal(expectedValue, value);
            Assert.Equal(expectedMatch, match);
        }


        [Theory(DisplayName = "Test of the ParserUtils TryParseDouble function")]
        [InlineData("1", 1, true)]
        [InlineData("0", 0, true)]
        [InlineData("-1", -1, true)]
        [InlineData("-10", -10, true)]
        [InlineData("22", 22, true)]
        [InlineData("-0.0000005", -0.0000005, true)]
        [InlineData("0.858", 0.858, true)]
        [InlineData("-99998.1234", -99998.1234, true)]
        [InlineData("98755142.2111137", 98755142.2111137, true)]
        [InlineData("", 0, false)]
        [InlineData(" ", 0, false)]
        [InlineData(null, 0, false)]
        [InlineData("a", 0, false)]
        [InlineData("1x1", 0, false)]
        [InlineData("0.0x", 0, false)]
        [InlineData("-22.5f4", 0, false)]
        public void TryParseDoubleTest(string givenValue, double expectedValue, bool expectedMatch)
        {
            bool match = ParserUtils.TryParseDouble(givenValue, out var value);
            Assert.Equal(expectedValue, value);
            Assert.Equal(expectedMatch, match);
        }


        [Fact(DisplayName = "Test of several numerical Parse and TryParse functions for their minimum values")]
        public void ParseMinTest()
        {
            bool match;

            match = ParserUtils.TryParseDecimal("-79228162514264337593543950335", out var dValue);
            Assert.Equal(decimal.MinValue, dValue);
            Assert.True(match);

            match = ParserUtils.TryParseUlong("0", out var uValue);
            Assert.Equal(ulong.MinValue, uValue);
            Assert.True(match);

            match = ParserUtils.TryParseLong("-9223372036854775808", out var lValue);
            Assert.Equal(long.MinValue, lValue);
            Assert.True(match);

            match = ParserUtils.TryParseUint("0", out var uiValue);
            Assert.Equal(uint.MinValue, uiValue);
            Assert.True(match);

            match = ParserUtils.TryParseInt("-2147483648", out var iValue);
            Assert.Equal(int.MinValue, iValue);
            Assert.True(match);

            iValue = ParserUtils.ParseInt("-2147483648");
            Assert.Equal(int.MinValue, iValue);

            match = ParserUtils.TryParseFloat("-3.40282347E+38", out var fValue);
            Assert.Equal(float.MinValue, fValue);
            Assert.True(match);

            fValue = ParserUtils.ParseFloat("-3.40282347E+38");
            Assert.Equal(float.MinValue, fValue);

            match = ParserUtils.TryParseDouble("-1.7976931348623157E+308", out var dbValue);
            Assert.Equal(double.MinValue, dbValue);
            Assert.True(match);
        }

        [Fact(DisplayName = "Test of several numerical Parse and TryParse functions for their maximum values")]
        public void ParseMaxTest()
        {
            bool match;

            match = ParserUtils.TryParseDecimal("79228162514264337593543950335", out var dValue);
            Assert.Equal(decimal.MaxValue, dValue);
            Assert.True(match);

            match = ParserUtils.TryParseUlong("18446744073709551615", out var uValue);
            Assert.Equal(ulong.MaxValue, uValue);
            Assert.True(match);

            match = ParserUtils.TryParseLong("9223372036854775807", out var lValue);
            Assert.Equal(long.MaxValue, lValue);
            Assert.True(match);

            match = ParserUtils.TryParseUint("4294967295", out var uiValue);
            Assert.Equal(uint.MaxValue, uiValue);
            Assert.True(match);

            match = ParserUtils.TryParseInt("2147483647", out var iValue);
            Assert.Equal(int.MaxValue, iValue);
            Assert.True(match);

            iValue = ParserUtils.ParseInt("2147483647");
            Assert.Equal(int.MaxValue, iValue);

            match = ParserUtils.TryParseFloat("3.40282347E+38", out var fValue);
            Assert.Equal(float.MaxValue, fValue);
            Assert.True(match);

            fValue = ParserUtils.ParseFloat("3.40282347E+38");
            Assert.Equal(float.MaxValue, fValue);

            match = ParserUtils.TryParseDouble("1.7976931348623157E+308", out var dbValue);
            Assert.Equal(double.MaxValue, dbValue);
            Assert.True(match);
        }



    }
}
