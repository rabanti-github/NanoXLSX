using System.Xml;
using NanoXLSX.Internal;
using Xunit;

namespace NanoXLSX.Test.Writer_Reader.Misc
{
    public class ReaderUtilsTests
    {
        [Theory(DisplayName = "Test of the GetAttribute Method when a value is returned on an existing attribute")]
        [InlineData("attr", "")]
        [InlineData("a", " ")]
        [InlineData("att", ".")]
        [InlineData("x", "0")]
        [InlineData("attr", "test")]
        [InlineData("at.tr", "123456789")]
        [InlineData("at-tr", " ")]
        [InlineData("at_tr", "at_tr")]
        [InlineData("__0014", "0000")]
        public void GetAttributeTest(string attributeName, string attributeValue)
        {
            // Arrange
            string xml = "<root " + attributeName + "='" + attributeValue + "'></root>";
            XmlDocument doc = new XmlDocument();
            doc.LoadXml(xml);
            XmlNode root = doc.DocumentElement;

            // Act
            string result = ReaderUtils.GetAttribute(root, attributeName);

            // Assert
            Assert.Equal(attributeValue, result);
        }

        [Fact(DisplayName = "Test of the GetAttribute Method when attribute does not exist (fallback value returned)")]
        public void GetAttributeFallbackTest()
        {
            // Arrange
            string xml = "<root></root>";
            XmlDocument doc = new XmlDocument();
            doc.LoadXml(xml);
            XmlNode root = doc.DocumentElement;

            // Act
            string result = ReaderUtils.GetAttribute(root, "nonexistent", "fallback");

            // Assert
            Assert.Equal("fallback", result);
        }

        [Fact(DisplayName = "Test of the GetAttribute Method when node has no attributes (e.g., a text node) returns fallback value")]
        public void GetAttributeNoAttributesTest()
        {
            // Arrange: Create a text node which does not support attributes (its Attributes property is null)
            XmlDocument doc = new XmlDocument();
            XmlNode textNode = doc.CreateTextNode("sample text");

            // Act
            string result = ReaderUtils.GetAttribute(textNode, "anyAttribute", "fallback");

            // Assert
            Assert.Equal("fallback", result);
        }

    }
}
