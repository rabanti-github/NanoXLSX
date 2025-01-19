using System;
using System.Collections.Generic;
using System.Linq;
using System.Security;
using System.Text;
using System.Threading.Tasks;
using NanoXLSX.Utils;
using Xunit;

namespace NanoXLSX.Core.Test.Utils
{
    public class ComparatorsTest
    {
        [Theory(DisplayName = "Test of the comparator function CompareSecureStrings")]
        [InlineData("", "", true)]
        [InlineData(" ", " ", true)]
        [InlineData("a", "a", true)]
        [InlineData("12345678", "12345678", true)]
        [InlineData("@à#", "@à#", true)]
        [InlineData("a", "A", false)]
        [InlineData("", " ", false)]
        [InlineData("123", "1234", false)]
        [InlineData("...", ".,.", false)]
        [InlineData(null, null, true)]
        [InlineData(null, "", true)] // Exception
        [InlineData(null, " ", false)]
        [InlineData(null, "ABC", false)]
        public void CompareSecureStringsTest(string plaintext1, string plaintext2, bool excpectedEqual)
        {
            SecureString sstr1 = GetSecureString(plaintext1);
            SecureString sstr2 = GetSecureString(plaintext2);
            bool isEqual = Comparators.CompareSecureStrings(sstr1, sstr2);
            Assert.Equal(excpectedEqual, isEqual);
            isEqual = Comparators.CompareSecureStrings(sstr2, sstr1); // reverse
            Assert.Equal(excpectedEqual, isEqual);
        }

        [Theory(DisplayName = "Test of the comparator function CompareDiomensions")]
        [InlineData(15.2f, 15.3f, -1)]
        [InlineData(15.3f, 15.2f, 1)]
        [InlineData(0.0002f, 0.0003f, -1)]
        [InlineData(0.0003f, 0.0002f, 1)]
        [InlineData(0.0002f, 0.0002f, 0)]
        [InlineData(1f, 2f, -1)]
        [InlineData(2f, 1f, 1)]
        [InlineData(-1f, 2f, -1)]
        [InlineData(-2f, -1f, -1)]
        [InlineData(-1f, -2f, 1)]
        [InlineData(0f, 0f, 0)]
        [InlineData(null, 15.3f, -1)]
        [InlineData(15.3f,null, 1)]
        [InlineData(null, null, 0)]
        public void CompareDimensionsTest(float? dimension1, float? dimension2, int expectedResult)
        {
            Assert.Equal(expectedResult, Comparators.CompareDimensions(dimension1, dimension2));
        }


        private static SecureString GetSecureString(string plainText)
        {
            if (plainText == null)
            {
                return null;
            }
            SecureString sstr = new SecureString();
            foreach (char c in plainText)
            {
                sstr.AppendChar(c);
            }
            return sstr;
        }
    }
}
