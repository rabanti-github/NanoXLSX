using System.Collections.Generic;
using NanoXLSX.Utils.Xml;
using Xunit;

namespace NanoXLSX.Core.Test.Utils.Xml
{
    public class XmlAttributeTest
    {
        [Theory(DisplayName = "CreateXmlAttributeTest: Should initialize properties correctly")]
        [InlineData("id", "123", "ns")]
        [InlineData("name", "value", "")]
        public void CreateXmlAttributeTest(string name, string value, string prefix)
        {
            var attribute = XmlAttribute.CreateAttribute(name, value, prefix);

            Assert.Equal(name, attribute.Name);
            Assert.Equal(value, attribute.Value);
            Assert.Equal(prefix, attribute.Prefix);
            Assert.Equal(!string.IsNullOrEmpty(prefix), attribute.HasPrefix);
        }

        [Theory(DisplayName = "CreateEmptyAttributeTest: Should create attribute with empty value")]
        [InlineData("empty", "")]
        [InlineData("test", "")]
        public void CreateEmptyAttributeTest(string name, string expectedValue)
        {
            var attribute = XmlAttribute.CreateEmptyAttribute(name);

            Assert.Equal(name, attribute.Name);
            Assert.Equal(expectedValue, attribute.Value);
            Assert.Equal("", attribute.Prefix);
            Assert.False(attribute.HasPrefix);
        }

        [Theory(DisplayName = "EqualsTest: Two attributes with same properties should be equal")]
        [InlineData("id", "123", "ns")]
        [InlineData("name", "value", "")]
        public void EqualsTest(string name, string value, string prefix)
        {
            var attribute1 = XmlAttribute.CreateAttribute(name, value, prefix);
            var attribute2 = XmlAttribute.CreateAttribute(name, value, prefix);

            Assert.True(attribute1.Equals(attribute2));
        }

        [Fact(DisplayName = "NotEqualsTest: Attributes with different properties should not be equal")]
        public void NotEqualsTest()
        {
            var attribute1 = XmlAttribute.CreateAttribute("id", "123", "ns");
            var attribute2 = XmlAttribute.CreateAttribute("id", "456", "ns");

            Assert.False(attribute1.Equals(attribute2));
        }

        [Theory(DisplayName = "GetHashCodeTest: Equal attributes should have the same hash code")]
        [InlineData("id", "123", "ns")]
        [InlineData("name", "value", "")]
        public void GetHashCodeTest(string name, string value, string prefix)
        {
            var attribute1 = XmlAttribute.CreateAttribute(name, value, prefix);
            var attribute2 = XmlAttribute.CreateAttribute(name, value, prefix);
            var hash1 = attribute1.GetHashCode();
            var hash2 = attribute2.GetHashCode();

            Assert.Equal(hash1, hash2);
        }

        [Fact(DisplayName = "FindAttribute - Matching when exactly one matching attribute is passed")]
        public void FindAttributeTest_OneMatching()
        {
            var attribute = XmlAttribute.CreateAttribute("test", "value");
            var attributes = new HashSet<XmlAttribute> { attribute };
            var result = XmlAttribute.FindAttribute("test", attributes);

            Assert.NotNull(result);
            Assert.Equal(attribute, result);
        }

        [Fact(DisplayName = "FindAttribute - Matching when a HashSet with multiple attributes contains a matching one")]
        public void FindAttributeTest_MatchingInSet()
        {
            var matchingAttribute = XmlAttribute.CreateAttribute("match", "value");
            var attributes = new HashSet<XmlAttribute>
            {
                XmlAttribute.CreateAttribute("other", "value"),
                matchingAttribute,
                XmlAttribute.CreateAttribute("another", "value")
            };
            var result = XmlAttribute.FindAttribute("match", attributes);

            Assert.NotNull(result);
            Assert.Equal(matchingAttribute, result);
        }

        [Fact(DisplayName = "FindAttribute - Non-matching when null is passed as HashSet")]
        public void FindAttributeTest_NullSet()
        {
            var result = XmlAttribute.FindAttribute("test", null);

            Assert.Null(result);
        }

        [Fact(DisplayName = "FindAttribute - Non-matching when an empty HashSet is passed")]
        public void FindAttributeTest_EmptySet()
        {
            var attributes = new HashSet<XmlAttribute>();
            var result = XmlAttribute.FindAttribute("test", attributes);

            Assert.Null(result);
        }

        [Fact(DisplayName = "FindAttribute - Non-matching when no attribute in the HashSet matches the name")]
        public void FindAttributeTest_NoMatch()
        {
            var attributes = new HashSet<XmlAttribute>
            {
                XmlAttribute.CreateAttribute("other", "value"),
                XmlAttribute.CreateAttribute("another", "value")
            };
            var result = XmlAttribute.FindAttribute("test", attributes);

            Assert.Null(result);
        }

    }
}
