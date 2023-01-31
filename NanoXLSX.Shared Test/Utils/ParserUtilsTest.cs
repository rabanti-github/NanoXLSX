using NanoXLSX.Shared.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xunit;

namespace NanoXLSX.Shared_Test.Utils
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
        [InlineData("true", 1)]
        [InlineData("TRUE", 1)]
        [InlineData("True", 1)]
        public void ParseBinaryBoolTest(String givenValue, int expectedValue)
        {
            int value = ParserUtils.ParseBinaryBool(givenValue);
            Assert.Equal(expectedValue, value);
        }
    }
}
