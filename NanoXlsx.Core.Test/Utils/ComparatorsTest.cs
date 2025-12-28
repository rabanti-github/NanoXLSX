using System.Security;
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
        public void CompareSecureStringsTest(string plainText1, string plainText2, bool expectedEqual)
        {
            SecureString sstr1 = GetSecureString(plainText1);
            SecureString sstr2 = GetSecureString(plainText2);
            bool isEqual = Comparators.CompareSecureStrings(sstr1, sstr2);
            Assert.Equal(expectedEqual, isEqual);
            isEqual = Comparators.CompareSecureStrings(sstr2, sstr1); // reverse
            Assert.Equal(expectedEqual, isEqual);
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
        [InlineData(15.3f, null, 1)]
        [InlineData(null, null, 0)]
        public void CompareDimensionsTest(float? dimension1, float? dimension2, int expectedResult)
        {
            Assert.Equal(expectedResult, Comparators.CompareDimensions(dimension1, dimension2));
        }

        [Theory(DisplayName = "Test of IsZero comparator (double)")]
        [InlineData(0.0, true)]
        [InlineData(-0.0, true)]
        [InlineData(1e-15, true)]
        [InlineData(-1e-15, true)]
        [InlineData(1e-13, true)]
        [InlineData(-1e-13, true)]
        [InlineData(1e-11, false)]
        [InlineData(-1e-11, false)]
        [InlineData(1.0, false)]
        [InlineData(-1.0, false)]
        [InlineData(double.Epsilon, true)]
        [InlineData(double.NaN, false)]
        [InlineData(double.PositiveInfinity, false)]
        [InlineData(double.NegativeInfinity, false)]
        public void IsZero_Double_Test(double value, bool expected)
        {
            Assert.Equal(expected, Comparators.IsZero(value));
        }

        [Theory(DisplayName = "Test of IsZero comparator (float)")]
        [InlineData(0.0f, true)]
        [InlineData(-0.0f, true)]
        [InlineData(1e-8f, true)]
        [InlineData(-1e-8f, true)]
        [InlineData(1e-6f, true)]
        [InlineData(-1e-6f, true)]
        [InlineData(1e-5f, false)]
        [InlineData(-1e-5f, false)]
        [InlineData(1.0f, false)]
        [InlineData(-1.0f, false)]
        [InlineData(float.Epsilon, true)]
        [InlineData(float.NaN, false)]
        [InlineData(float.PositiveInfinity, false)]
        [InlineData(float.NegativeInfinity, false)]
        public void IsZero_Float_Test(float value, bool expected)
        {
            Assert.Equal(expected, Comparators.IsZero(value));
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
