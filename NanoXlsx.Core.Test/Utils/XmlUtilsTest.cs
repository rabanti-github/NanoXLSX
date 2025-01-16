using NanoXLSX.Utils;
using Xunit;

namespace NanoXLSX.Test.Core.UtilsTest
{
    public class XmlUtilsTest
    {
        [Theory(DisplayName = "Test of the EscapeXmlChars function")]
        [InlineData(null, "")]
        [InlineData("", "")]
        [InlineData(" ", " ")]
        [InlineData("This is a test & string", "This is a test &amp; string")]
        [InlineData("This is a <tag>", "This is a &lt;tag&gt;")]
        [InlineData("This is a >tag<", "This is a &gt;tag&lt;")]
        [InlineData("This is a \"quoted\" text", "This is a \"quoted\" text")]
        public void EscapeXmlCharsTest(string input, string expectedOutput)
        {
            string result = XmlUtils.EscapeXmlChars(input);
            Assert.Equal(expectedOutput, result);
        }

        [Theory(DisplayName = "Test of the EscapeXmlChars function with special characters")]
        [InlineData("\x01", " ")]
        [InlineData("\x02", " ")]
        [InlineData("\x10", " ")]
        [InlineData("\x1F", " ")]
        //[InlineData("\x7F", " ")] // valid in XML v1.0
        public void EscapeXmlCharsSpecialCharactersTest(string input, string expectedOutput)
        {
            string result = XmlUtils.EscapeXmlChars(input);
            Assert.Equal(expectedOutput, result);
        }

        [Theory(DisplayName = "Test of the EscapeXmlAttributeChars function")]
        [InlineData("This is a test & string", "This is a test &amp; string")]
        [InlineData("This is a <tag>", "This is a &lt;tag&gt;")]
        [InlineData("This is a >tag<", "This is a &gt;tag&lt;")]
        [InlineData("This is a \"quoted\" text", "This is a &quot;quoted&quot; text")]
        public void EscapeXmlAttributeCharsTest(string input, string expectedOutput)
        {
            string result = XmlUtils.EscapeXmlAttributeChars(input);
            Assert.Equal(expectedOutput, result);
        }

        [Theory(DisplayName = "Test of the EscapeXmlAttributeChars function with special characters")]
        [InlineData("\x01", " ")]        // Invalid control character
        [InlineData("\x02", " ")]        // Invalid control character
        [InlineData("\x10", " ")]        // Invalid control character
        [InlineData("\x1F", " ")]        // Invalid control character
        [InlineData("\xD800", "\xFFFD\xFFFD")]    // Invalid high surrogate (deliberate invalid input)
        [InlineData("\uDBFF", "\xFFFD\xFFFD")]    // Invalid high surrogate (deliberate invalid input)
        [InlineData("\uDC00", "\xFFFD\xFFFD")]    // Invalid low surrogate (deliberate invalid input)
        [InlineData("\uDFFF", "\xFFFD\xFFFD")]    // Invalid low surrogate (deliberate invalid input)
        [InlineData("\uD835\uDC00", "&#x1D400;")] // Valid surrogate pair representing a Unicode character
        [InlineData("\uFFFE", " ")]      // Invalid character
        [InlineData("\uFFFF", " ")]      // Invalid character
        public void EscapeXmlAttributeCharsSpecialCharactersTest(string input, string expectedOutput)
        {
            string result = XmlUtils.EscapeXmlAttributeChars(input);
            Assert.Equal(expectedOutput, result);
        }
    }
}
