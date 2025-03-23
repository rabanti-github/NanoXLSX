using NanoXLSX.Utils.Xml;
using Xunit;

namespace NanoXLSX.Test.Core.UtilsTest
{
    public class XmlUtilsTest
    {
       [Theory(DisplayName = "Test of the SanitizeXmlValue function")]
       [InlineData(null, "")]
       [InlineData("", "")]
       [InlineData(" ", " ")]
       [InlineData("This is a test & string", "This is a test & string")] // not escaped since handled by writer
       [InlineData("This is a <tag>", "This is a <tag>")] // not escaped since handled by writer
       [InlineData("This is a >tag<", "This is a >tag<")] // not escaped since handled by writer
       [InlineData("This is a \"quoted\" text", "This is a \"quoted\" text")]
       public void SanitizeXmlValueTest(string input, string expectedOutput)
       {
           string result = XmlUtils.SanitizeXmlValue(input);
           Assert.Equal(expectedOutput, result);
       }

        [Theory(DisplayName = "Test of the SanitizeXmlValue function with special characters")]
        [InlineData("\x01", " ")]
        [InlineData("\x02", " ")]
        [InlineData("\x10", " ")]
        [InlineData("\x1F", " ")]
        //[InlineData("\x7F", " ")] // valid in XML v1.0
        public void SanitizeXmlValueSpecialCharactersTest(string input, string expectedOutput)
        {
            string result = XmlUtils.SanitizeXmlValue(input);
            Assert.Equal(expectedOutput, result);
        }
    }
}
