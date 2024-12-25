using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xunit;

namespace NanoXLSX.Core.Test.Misc
{
    public class LegacyPasswordTest
    {
        [Theory(DisplayName = "Test of the GeneratePasswordHash function (legacy)")]
        [InlineData("x", "CEBA")]
        [InlineData("Test@1-2,3!", "F767")]
        [InlineData(" ", "CE0A")]
        [InlineData("", "")]
        [InlineData(null, "")]
        public void GeneratePasswordHashTest(string givenVPassword, string expectedHash)
        {
            string hash = LegacyPassword.GenerateLegacyPasswordHash(givenVPassword);
            Assert.Equal(expectedHash, hash);
        }
    }
}
