using NanoXLSX.Shared.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xunit;

namespace NanoXLSX.Shared_Test.Utils
{
    public class ValidatorsTest
    {
        [Theory(DisplayName = "Test of the successful Validator function ValidateColor")]
        [InlineData("000000", false, true)]
        [InlineData("000000", false, false)]
        [InlineData("00AACC", false, true)]
        [InlineData("00AACC", false, false)]
        [InlineData("FFFFFF", false, true)]
        [InlineData("FFFFFF", false, false)]
        [InlineData("00000000", true, true)]
        [InlineData("00000000", true, false)]
        [InlineData("FF000000", true, true)]
        [InlineData("FF000000", true, false)]
        [InlineData("00AACC00", true, true)]
        [InlineData("00AACC00", true, false)]
        [InlineData("", true, true)]
        [InlineData("", false, true)]
        [InlineData(null, true, true)]
        [InlineData(null, false, true)]
        public void ValidateColorTest(string givenHexCode, bool givenUseAlpha, bool givenAllowEmpty)
        {
                Validators.ValidateColor(givenHexCode, givenUseAlpha, givenAllowEmpty);
                Assert.True(true);
        }

        [Theory(DisplayName = "Test of the failing Validator function ValidateColor")]
        [InlineData("000000", true, true)]
        [InlineData("FFFFFF", true, true)]
        [InlineData("0ACFD", true, true)]
        [InlineData("00000000", false, true)]
        [InlineData("00FFFFFF", false, true)]
        [InlineData("000ACFD", false, true)]
        [InlineData("FF000000", false, true)]
        [InlineData("FFFFFFFF", false, true)]
        [InlineData("FF0ACFD", false, true)]
        [InlineData("AA", false, true)]
        [InlineData("CCC", false, true)]
        [InlineData("DDDD", false, true)]
        [InlineData("001122", true, true)]
        [InlineData("X", false, true)]
        [InlineData("AAX022", false, true)]
        [InlineData(" ", false, true)]
        [InlineData("0 0000", false, true)]
        [InlineData("", false, false)]
        [InlineData(null, false, false)]

        public void ValidateColorFailTest(string givenHexCode, bool givenUseAlpha, bool givenAllowEmpty)
        {
                Assert.Throws<NanoXLSX.Shared.Exceptions.StyleException>(() => Validators.ValidateColor(givenHexCode, givenUseAlpha, givenAllowEmpty));
        }
    }
}
