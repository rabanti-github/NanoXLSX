using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NanoXLSX.Utils.Xml;
using Xunit;

namespace NanoXLSX.Core.Test.UtilsTest
{
    public class XmlAttributeTest
    {
        [Theory(DisplayName = "CreateXmlAttributeTest: Should initialize properties correctly")]
        [InlineData("id", "123", "ns")]
        [InlineData("name", "value", "")]
        public void CreateXmlAttributeTest(string name, string value, string prefix)
        {
            XmlAttribute attribute = XmlAttribute.CreateAttribute(name, value, prefix);

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
            XmlAttribute attribute = XmlAttribute.CreateEmptyAttribute(name);

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
            XmlAttribute attribute1 = XmlAttribute.CreateAttribute(name, value, prefix);
            XmlAttribute attribute2 = XmlAttribute.CreateAttribute(name, value, prefix);

            Assert.True(attribute1.Equals(attribute2));
        }

        [Fact(DisplayName = "NotEqualsTest: Attributes with different properties should not be equal")]
        public void NotEqualsTest()
        {
            XmlAttribute attribute1 = XmlAttribute.CreateAttribute("id", "123", "ns");
            XmlAttribute attribute2 = XmlAttribute.CreateAttribute("id", "456", "ns");

            Assert.False(attribute1.Equals(attribute2));
        }

        [Theory(DisplayName = "GetHashCodeTest: Equal attributes should have the same hash code")]
        [InlineData("id", "123", "ns")]
        [InlineData("name", "value", "")]
        public void GetHashCodeTest(string name, string value, string prefix)
        {
            XmlAttribute attribute1 = XmlAttribute.CreateAttribute(name, value, prefix);
            XmlAttribute attribute2 = XmlAttribute.CreateAttribute(name, value, prefix);

            int hash1 = attribute1.GetHashCode();
            int hash2 = attribute2.GetHashCode();

            Assert.Equal(hash1, hash2);
        }
    }
}
