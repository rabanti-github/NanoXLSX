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

        [Theory(DisplayName = "Test of the ParserUtils ToString function")]
        [InlineData(-10, "-10")]
        [InlineData(0, "0")]
        [InlineData(1, "1")]
        [InlineData(100, "100")]
        public void ToStringTest(int givenValue, string expectedValue)
        {
            string value = ParserUtils.ToString(givenValue);
            Assert.Equal(expectedValue, value);
        }
    }
}
