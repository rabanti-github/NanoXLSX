using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
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

        [Theory(DisplayName = "Test of the GetChildNode Method ignoring case of the node name")]
        [InlineData("child", "Child")]
        [InlineData("Child", "child")]
        [InlineData("NODE", "node")]
        [InlineData("_nOdE", "_node")]
        [InlineData("chilD-0", "Child-0")]
        public void GetChildNodeTest(string nodeName, string searchName)
        {
            // Arrange
            string xml = "<root><" + nodeName + ">content</" + nodeName + "></root>";
            XmlDocument doc = new XmlDocument();
            doc.LoadXml(xml);
            XmlNode root = doc.DocumentElement;

            // Act
            XmlNode child = ReaderUtils.GetChildNode(root,  searchName);

            // Assert
            Assert.NotNull(child);
            Assert.Equal(nodeName, child.Name);
        }

        [Fact(DisplayName = "Test of the GetChildNode Method when node is null returns null")]
        public void GetChildNodeNodeNullTest()
        {
            // Act
            XmlNode child = ReaderUtils.GetChildNode(null, "child");

            // Assert
            Assert.Null(child);
        }


        [Theory(DisplayName = "Test of the GetAttributeOfChild Method when child and attribute exist")]
        [InlineData("child", "att", "content")]
        [InlineData("_Child", "att0", "0001")]
        [InlineData("child-001", "__att__", "true")]
        [InlineData("child.a", "ATT0001", " ")]
        [InlineData("_Child", "att0", "")]
        public void GetAttributeOfChildTest(string nodeName, string attributeName, string attributeValue)
        {
            // Arrange
            string xml = "<root><" + nodeName + " " + attributeName + "='" + attributeValue + "' /></root>";
            XmlDocument doc = new XmlDocument();
            doc.LoadXml(xml);
            XmlNode root = doc.DocumentElement;

            // Act
            bool found = ReaderUtils.GetAttributeOfChild(root, nodeName, attributeName, out string output);

            // Assert
            Assert.True(found);
            Assert.Equal(attributeValue, output);
        }

        [Fact(DisplayName = "Test of the GetAttributeOfChild Method when child does not exist")]
        public void GetAttributeOfChildMissingTest()
        {
            // Arrange
            string xml = "<root><anotherchild attr='value' /></root>";
            XmlDocument doc = new XmlDocument();
            doc.LoadXml(xml);
            XmlNode root = doc.DocumentElement;

            // Act
            bool found = ReaderUtils.GetAttributeOfChild(root, "child", "attr", out string output);

            // Assert
            Assert.False(found);
            Assert.Null(output);
        }

        [Theory(DisplayName = "Test of the IsNode Method for a matching node (ignoring case)")]
        [InlineData("child", "Child")]
        [InlineData("Child", "child")]
        [InlineData("NODE", "node")]
        [InlineData("_nOdE", "_node")]
        [InlineData("chilD-0", "Child-0")]
        public void IsNodeTest(string nodeName, string searchName)
        {
            // Arrange
            string xml = "<" + nodeName + "></" + nodeName + ">";
            XmlDocument doc = new XmlDocument();
            doc.LoadXml(xml);
            XmlNode node = doc.DocumentElement;

            // Act
            bool result = ReaderUtils.IsNode(node, searchName);

            // Assert
            Assert.True(result);
        }

        [Fact(DisplayName = "Test of the DiscoverPrefix Method when element has a prefix")]
        public void DiscoverPrefixTest()
        {
            // Arrange
            string xml = "<p:root xmlns:p='http://example.com'></p:root>";
            XmlDocument doc = new XmlDocument();
            doc.LoadXml(xml);

            // Act
            string prefix = ReaderUtils.DiscoverPrefix(doc, "root");

            // Assert
            Assert.Equal("p", prefix);
        }

        [Fact(DisplayName = "Test of the DiscoverPrefix Method when no element matches")]
        public void DiscoverPrefixMissingTest()
        {
            // Arrange
            string xml = "<p:root xmlns:p='http://example.com'></p:root>";
            XmlDocument doc = new XmlDocument();
            doc.LoadXml(xml);

            // Act
            string prefix = ReaderUtils.DiscoverPrefix(doc, "nonexistent");

            // Assert
            Assert.Equal(string.Empty, prefix);
        }

        [Fact(DisplayName = "Test of the DiscoverPrefix Method when no no prefix was defined")]
        public void DiscoverPrefixMissingTest2()
        {
            // Arrange
            string xml = "<root>content</root>";
            XmlDocument doc = new XmlDocument();
            doc.LoadXml(xml);

            // Act
            string prefix = ReaderUtils.DiscoverPrefix(doc, "root");

            // Assert
            Assert.Equal(string.Empty, prefix);
        }

        [Fact(DisplayName = "Test of the GetElementsByTagName Method without prefix")]
        public void GetElementsByTagNameNoPrefixTest()
        {
            // Arrange
            string xml = "<root><child>one</child><child>two</child></root>";
            XmlDocument doc = new XmlDocument();
            doc.LoadXml(xml);

            // Act
            XmlNodeList nodes = ReaderUtils.GetElementsByTagName(doc, "child", "");

            // Assert
            Assert.NotNull(nodes);
            Assert.Equal(2, nodes.Count);
        }

        [Fact(DisplayName = "Test of the GetElementsByTagName Method with prefix")]
        public void GetElementsByTagNameWithPrefixTest()
        {
            // Arrange
            string xml = "<p:root xmlns:p='http://example.com'><p:child>one</p:child><p:child>two</p:child></p:root>";
            XmlDocument doc = new XmlDocument();
            doc.LoadXml(xml);

            // Act
            XmlNodeList nodes = ReaderUtils.GetElementsByTagName(doc, "child", "p");

            // Assert
            Assert.NotNull(nodes);
            Assert.Equal(2, nodes.Count);
        }
    }
}
