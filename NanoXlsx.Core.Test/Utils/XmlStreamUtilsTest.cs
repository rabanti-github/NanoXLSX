using System.IO;
using System.Xml;
using NanoXLSX.Utils.Xml;
using Xunit;

namespace NanoXLSX.Core.Test.UtilsTest
{
    public class XmlStreamUtilsTest
    {
        #region CreateSettings

        [Fact(DisplayName = "CreateSettings - IgnoreWhitespace should be true")]
        public void CreateSettings_IgnoreWhitespaceShouldBeTrue()
        {
            XmlReaderSettings settings = XmlStreamUtils.CreateSettings();
            Assert.True(settings.IgnoreWhitespace);
        }

        [Fact(DisplayName = "CreateSettings - IgnoreComments should be true")]
        public void CreateSettings_IgnoreCommentsShouldBeTrue()
        {
            XmlReaderSettings settings = XmlStreamUtils.CreateSettings();
            Assert.True(settings.IgnoreComments);
        }

        [Fact(DisplayName = "CreateSettings - IgnoreProcessingInstructions should be true")]
        public void CreateSettings_IgnoreProcessingInstructionsShouldBeTrue()
        {
            XmlReaderSettings settings = XmlStreamUtils.CreateSettings();
            Assert.True(settings.IgnoreProcessingInstructions);
        }

        #endregion

        #region IsElement

        [Theory(DisplayName = "IsElement - Should return true when positioned on a matching start element (case-insensitive)")]
        [InlineData("root", "root")]
        [InlineData("Root", "root")]
        [InlineData("root", "Root")]
        [InlineData("ROOT", "root")]
        [InlineData("node", "NODE")]
        public void IsElement_Match(string elementName, string localName)
        {
            using (XmlReader reader = CreateReaderAtFirstElement($"<{elementName}/>"))
            {
                Assert.True(XmlStreamUtils.IsElement(reader, localName));
            }
        }

        [Theory(DisplayName = "IsElement - Should return false when the local name does not match")]
        [InlineData("root", "node")]
        [InlineData("child", "root")]
        [InlineData("element", "elements")]
        public void IsElement_NoMatch(string elementName, string localName)
        {
            using (XmlReader reader = CreateReaderAtFirstElement($"<{elementName}/>"))
            {
                Assert.False(XmlStreamUtils.IsElement(reader, localName));
            }
        }

        [Fact(DisplayName = "IsElement - Should return false when positioned on an end element")]
        public void IsElement_EndElement()
        {
            using (XmlReader reader = XmlReader.Create(new StringReader("<root></root>"), XmlStreamUtils.CreateSettings()))
            {
                reader.Read(); // positioned on <root> start element
                reader.Read(); // positioned on </root> end element
                Assert.Equal(XmlNodeType.EndElement, reader.NodeType);
                Assert.False(XmlStreamUtils.IsElement(reader, "root"));
            }
        }

        [Fact(DisplayName = "IsElement - Should strip namespace prefix and match only local name")]
        public void IsElement_NamespacePrefix()
        {
            using (XmlReader reader = XmlReader.Create(new StringReader("<ns:root xmlns:ns=\"http://example.com\"/>"), XmlStreamUtils.CreateSettings()))
            {
                reader.Read();
                Assert.True(XmlStreamUtils.IsElement(reader, "root"));
            }
        }

        #endregion

        #region ReadElementText

        [Fact(DisplayName = "ReadElementText - Should return empty string for a self-closing (empty) element")]
        public void ReadElementText_EmptyElement()
        {
            using (XmlReader reader = CreateReaderAtFirstElement("<node/>"))
            {
                string result = XmlStreamUtils.ReadElementText(reader);
                Assert.Equal(string.Empty, result);
            }
        }

        [Fact(DisplayName = "ReadElementText - Should return empty string for an open/close element with no content")]
        public void ReadElementText_OpenCloseEmptyElement()
        {
            using (XmlReader reader = CreateReaderAtFirstElement("<node></node>"))
            {
                string result = XmlStreamUtils.ReadElementText(reader);
                Assert.Equal(string.Empty, result);
            }
        }

        [Theory(DisplayName = "ReadElementText - Should return the text content of a leaf element")]
        [InlineData("<node>hello</node>", "hello")]
        [InlineData("<node>123</node>", "123")]
        [InlineData("<node>  spaces  </node>", "  spaces  ")]
        [InlineData("<node>line1</node>", "line1")]
        public void ReadElementText_WithContent(string xml, string expected)
        {
            using (XmlReader reader = CreateReaderAtFirstElement(xml))
            {
                string result = XmlStreamUtils.ReadElementText(reader);
                Assert.Equal(expected, result);
            }
        }

        #endregion

        #region Helpers

        private static XmlReader CreateReaderAtFirstElement(string xml)
        {
            XmlReader reader = XmlReader.Create(new StringReader(xml), XmlStreamUtils.CreateSettings());
            reader.Read();
            return reader;
        }

        #endregion
    }
}
