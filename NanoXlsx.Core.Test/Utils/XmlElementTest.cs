using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NanoXLSX.Utils.Xml;
using Xunit;

namespace NanoXLSX.Core.Test.UtilsTest
{
    public class XmlElementTest
    {
        [Theory(DisplayName = "Constructor should correctly set Name and Prefix and leave properties null")]
        [InlineData("ElementName", "prefix")]
        [InlineData("ElementName", "")]
        [InlineData("AnotherElement", "somePrefix")]
        public void CreateXmlElementTest(string name, string prefix)
        {
            XmlElement element = XmlElement.CreateElement(name, prefix);
            // Assert: Check read-only properties and default state.
            Assert.Equal(name, element.Name);
            Assert.Equal(prefix, element.Prefix);
            Assert.Null(element.Children);
            Assert.Null(element.Attributes);
            Assert.Null(element.PrefixNameSpaceMap);
        }

        [Theory(DisplayName = "Prefix property should be get and set correctly")]
        [InlineData("initialPrefix", "newPrefix")]
        [InlineData("", "nonEmptyPrefix")]
        public void PrefixPropertyTest(string initialPrefix, string newPrefix)
        {
            XmlElement element = XmlElement.CreateElement("TestElement", initialPrefix);
            element.Prefix = newPrefix;
            Assert.Equal(newPrefix, element.Prefix);
        }

        [Theory(DisplayName = "InnerValue property should set value if non-empty; empty or null resets to null")]
        [InlineData("Some value", "Some value")]
        [InlineData("", null)]
        [InlineData(null, null)]
        public void InnerValuePropertyTest(string setValue, string expectedValue)
        {
            XmlElement element = XmlElement.CreateElement("TestElement");
            element.InnerValue = setValue;
            Assert.Equal(expectedValue, element.InnerValue);
        }

        [Fact(DisplayName = "Children property should be null when no children have been added")]
        public void ChildrenPropertyInitialTest()
        {
            XmlElement element = XmlElement.CreateElement("TestElement");

            Assert.Null(element.Children);
        }

        [Fact(DisplayName = "Attributes property should be null when no attributes have been added")]
        public void AttributesPropertyInitialTest()
        {
            XmlElement element = XmlElement.CreateElement("TestElement");
            Assert.Null(element.Attributes);
        }

        [Fact(DisplayName = "PrefixNameSpaceMap property should be null when not set")]
        public void PrefixNameSpaceMapPropertyInitialTest()
        {
            XmlElement element = XmlElement.CreateElement("TestElement");
            Assert.Null(element.PrefixNameSpaceMap);
        }

        [Theory(DisplayName = "AddAttribute(string, string, string) should add a single attribute correctly")]
        [InlineData("attr1", "value1", "prefix1")]
        [InlineData("attr2", "value2", "")]
        public void AddAttributeStringMethodTest(string name, string value, string prefix)
        {
            XmlElement element = XmlElement.CreateElement("TestElement");
            element.AddAttribute(name, value, prefix);
            Assert.NotNull(element.Attributes);
            Assert.Single(element.Attributes);
            // Get the attribute added (HashSet does not guarantee order, so we take the first)
            XmlAttribute attr = element.Attributes.First();
            Assert.Equal(name, attr.Name);
            Assert.Equal(value, attr.Value);
            Assert.Equal(prefix, attr.Prefix);
        }

        [Fact(DisplayName = "AddAttribute(XmlAttribute?) should add a valid attribute and ignore null values")]
        public void AddAttributeNullableAttributeTest()
        {
            XmlElement element = XmlElement.CreateElement("TestElement");
            XmlAttribute validAttribute = XmlAttribute.CreateAttribute("attrValid", "valueValid", "pfx");
            element.AddAttribute(validAttribute);
            XmlAttribute? nullAttribute = default;
            element.AddAttribute(nullAttribute);
            Assert.NotNull(element.Attributes);
            // Only the valid attribute should have been added.
            Assert.Single(element.Attributes);
            XmlAttribute attr = element.Attributes.First();
            Assert.Equal("attrValid", attr.Name);
            Assert.Equal("valueValid", attr.Value);
            Assert.Equal("pfx", attr.Prefix);
        }

        [Fact(DisplayName = "AddAttributes(IEnumerable<XmlAttribute>) should add multiple attributes, and ignore null/empty collections")]
        public void AddAttributesEnumerableTest()
        {
            XmlElement element = XmlElement.CreateElement("TestElement");
            List<XmlAttribute> attributesList = new List<XmlAttribute>
            {
                XmlAttribute.CreateAttribute("attrA", "valueA", "pfxA"),
                XmlAttribute.CreateAttribute("attrB", "valueB")
            };
            element.AddAttributes(attributesList);
            Assert.NotNull(element.Attributes);
            Assert.Equal(attributesList.Count, element.Attributes.Count);
            
            element.AddAttributes(new List<XmlAttribute>());
            Assert.Equal(attributesList.Count, element.Attributes.Count);
            element.AddAttributes(null);
            Assert.Equal(attributesList.Count, element.Attributes.Count);
        }

        [Theory(DisplayName = "AddNameSpaceAttribute should add namespace mapping and corresponding attribute when valid")]
        [InlineData("ns", "xmlns", "http://example.com/ns")]
        [InlineData("x", "xmlns", "http://example.org/x")]
        public void AddNameSpaceAttributeValidInputTest(string prefix, string rootNameSpace, string uri)
        {
            XmlElement element = XmlElement.CreateElement("TestElement", "t");
            element.AddNameSpaceAttribute(prefix, rootNameSpace, uri);

            Assert.NotNull(element.PrefixNameSpaceMap);
            Assert.True(element.PrefixNameSpaceMap.ContainsKey(prefix));
            Assert.Equal(uri, element.PrefixNameSpaceMap[prefix]);

            Assert.NotNull(element.Attributes);
            XmlAttribute nsAttribute = element.Attributes.FirstOrDefault(attr => attr.Name == prefix);
            Assert.Equal(uri, nsAttribute.Value);
            Assert.Equal(rootNameSpace, nsAttribute.Prefix);
        }

        [Theory(DisplayName = "AddNameSpaceAttribute should ignore empty prefix or URI")]
        [InlineData("", "xmlns", "http://example.com/ns")]
        [InlineData("ns", "xmlns", "")]
        [InlineData("", "xmlns", "")]
        public void AddNameSpaceAttributeInvalidInputTest(string prefix, string rootNameSpace, string uri)
        {
            XmlElement element = XmlElement.CreateElement("TestElement", "t");
            element.AddNameSpaceAttribute(prefix, rootNameSpace, uri);

            Assert.Null(element.PrefixNameSpaceMap);
            Assert.Null(element.Attributes);
        }

        [Theory(DisplayName = "AddDefaultXmlNameSpace should set the default XML namespace for the element")]
        [InlineData("http://example.com/default")]
        [InlineData("http://example.org/ns")]
        public void AddDefaultXmlNameSpaceTest(string defaultUri)
        {
            XmlElement element = XmlElement.CreateElement("TestElement");
            element.AddDefaultXmlNameSpace(defaultUri);
            System.Xml.XmlDocument doc = element.TransformToDocument();

            // Assert: When a default namespace is defined, the root element should have it set.
            Assert.NotNull(doc.DocumentElement);
            Assert.Equal("TestElement", doc.DocumentElement.LocalName);
            Assert.Equal(defaultUri, doc.DocumentElement.NamespaceURI);
        }

        [Theory(DisplayName = "AddChildElementWithAttribute should create a child with one attribute and add it to the parent's children")]
        [InlineData("ChildName", "attrName", "attrValue", "childPrefix", "attrPrefix")]
        public void AddChildElementWithAttributeTest(string childName, string attributeName, string attributeValue, string namePrefix, string attributePrefix)
        {
            XmlElement parent = XmlElement.CreateElement("Parent");
            XmlElement child = parent.AddChildElementWithAttribute(childName, attributeName, attributeValue, namePrefix, attributePrefix);

            Assert.NotNull(child);
            Assert.NotNull(parent.Children);
            Assert.Contains(child, parent.Children);

            Assert.NotNull(child.Attributes);
            Assert.Single(child.Attributes);
            XmlAttribute attr = child.Attributes.First();
            Assert.Equal(attributeName, attr.Name);
            Assert.Equal(attributeValue, attr.Value);
            Assert.Equal(attributePrefix, attr.Prefix);
        }

        [Theory(DisplayName = "AddChildElementWithValue should create a child with inner value when provided; returns null for empty inner value")]
        [InlineData("ChildName", "Inner Text", "childPrefix", true)]
        [InlineData("ChildName", "", "childPrefix", false)]
        public void AddChildElementWithValueTest(string childName, string innerValue, string prefix, bool shouldBeAdded)
        {
            XmlElement parent = XmlElement.CreateElement("Parent");
            XmlElement child = parent.AddChildElementWithValue(childName, innerValue, prefix);

            if (shouldBeAdded)
            {
                Assert.NotNull(child);
                Assert.NotNull(parent.Children);
                Assert.Contains(child, parent.Children);
                Assert.Equal(innerValue, child.InnerValue);
            }
            else
            {
                Assert.Null(child);
                Assert.Null(parent.Children);
            }
        }

        [Theory(DisplayName = "AddChildElement(string, string) should create and add a child element")]
        [InlineData("ChildName", "childPrefix")]
        [InlineData("AnotherChild", "")]
        public void AddChildElementStringOverloadTest(string childName, string prefix)
        {
            XmlElement parent = XmlElement.CreateElement("Parent");
            XmlElement child = parent.AddChildElement(childName, prefix);

            Assert.NotNull(child);
            Assert.NotNull(parent.Children);
            Assert.Contains(child, parent.Children);
            Assert.Equal(childName, child.Name);
            Assert.Equal(prefix, child.Prefix);
        }

        [Fact(DisplayName = "AddChildElement(XmlElement) should add a non-null child and ignore null")]
        public void AddChildElementXmlElementOverloadTest()
        {
            XmlElement parent = XmlElement.CreateElement("Parent");
            XmlElement child = XmlElement.CreateElement("Child", "c");

            parent.AddChildElement(child);

            Assert.NotNull(parent.Children);
            Assert.Contains(child, parent.Children);
            int countAfterValid = parent.Children.Count;

            parent.AddChildElement(null);
            Assert.Equal(countAfterValid, parent.Children.Count);
        }

        [Fact(DisplayName = "AddChildElements(IEnumerable<XmlElement>) should add multiple children and ignore null or empty collections")]
        public void AddChildElementsEnumerableTest()
        {
            XmlElement parent = XmlElement.CreateElement("Parent");
            XmlElement child1 = XmlElement.CreateElement("Child1");
            XmlElement child2 = XmlElement.CreateElement("Child2");
            List<XmlElement> childrenList = new List<XmlElement> { child1, child2 };

            parent.AddChildElements(childrenList);

            Assert.NotNull(parent.Children);
            Assert.Equal(childrenList.Count, parent.Children.Count);
            Assert.Contains(child1, parent.Children);
            Assert.Contains(child2, parent.Children);

            parent.AddChildElements(new List<XmlElement>());
            Assert.Equal(childrenList.Count, parent.Children.Count);

            parent.AddChildElements(null);
            Assert.Equal(childrenList.Count, parent.Children.Count);
        }

        [Theory(DisplayName = "CreateElement should instantiate an element with the given name and optional prefix")]
        [InlineData("TestElement", "prefix")]
        [InlineData("TestElement", "")]
        [InlineData("AnotherElement", "ns")]
        public void CreateElementTest(string name, string prefix)
        {
            XmlElement element = XmlElement.CreateElement(name, prefix);

            Assert.NotNull(element);
            Assert.Equal(name, element.Name);
            Assert.Equal(prefix, element.Prefix);
            Assert.Null(element.Attributes);
            Assert.Null(element.Children);
            Assert.Null(element.PrefixNameSpaceMap);
        }

        [Theory(DisplayName = "CreateElementWithAttribute should instantiate an element with one attribute")]
        [InlineData("ElementWithAttr", "attrName", "attrValue", "elemPrefix", "attrPrefix")]
        [InlineData("ElementWithAttr", "id", "123", "", "")]
        public void CreateElementWithAttributeTest(string name, string attributeName, string attributeValue, string namePrefix, string attributePrefix)
        {
            XmlElement element = XmlElement.CreateElementWithAttribute(name, attributeName, attributeValue, namePrefix, attributePrefix);

            Assert.NotNull(element);
            Assert.Equal(name, element.Name);
            Assert.Equal(namePrefix, element.Prefix);
            Assert.NotNull(element.Attributes);
            Assert.Single(element.Attributes);
            XmlAttribute attr = element.Attributes.First();
            Assert.Equal(attributeName, attr.Name);
            Assert.Equal(attributeValue, attr.Value);
            Assert.Equal(attributePrefix, attr.Prefix);
        }

        [Theory(DisplayName = "TransformToDocument should create an XmlDocument with correct hierarchy, attributes, and inner text, with and without default namespace")]
        [InlineData(true)]
        [InlineData(false)]
        public void TransformToDocumentTest(bool useDefaultNamespace)
        {
            XmlElement root = XmlElement.CreateElement("Root", "r");
            if (useDefaultNamespace)
            {
                // Set custom default namespace so that all elements use it.
                root.AddDefaultXmlNameSpace("http://example.com/ns");
            }
            else
            {
                // Set namespace via attribute (will be skipped for root creation)
                root.AddNameSpaceAttribute("xmlns", "", "http://example.com/ns");
            }
            root.AddAttribute("version", "1.0");

            // Create a child element with one attribute.
            // If using default namespace, create the child with an empty prefix to have the default applied.
            // Otherwise, use a specific prefix (e.g., "xmlns") as in the original test.
            XmlElement childWithAttr = useDefaultNamespace
                ? root.AddChildElementWithAttribute("Child", "id", "123", "", "")
                : root.AddChildElementWithAttribute("Child", "id", "123", "xmlns", "");
            childWithAttr.InnerValue = "ChildValue";

            System.Xml.XmlDocument doc = root.TransformToDocument();

            Assert.NotNull(doc.DocumentElement);
            Assert.Equal("Root", doc.DocumentElement.LocalName);
            string versionAttr = doc.DocumentElement.GetAttribute("version");
            Assert.Equal("1.0", versionAttr);

            Assert.True(doc.DocumentElement.ChildNodes.Count >= 1, "The root element should have at least one child element.");

            System.Xml.XmlElement childElement = doc.DocumentElement.ChildNodes
                .OfType<System.Xml.XmlElement>()
                .FirstOrDefault(e => e.LocalName == "Child");

            if (useDefaultNamespace)
            {
                // User defined default name space
                Assert.Equal("http://example.com/ns", childElement.NamespaceURI);
            }
            else
            {
                // Fall back to the general default name space
                Assert.Equal("http://www.w3.org/2000/xmlns/", childElement.NamespaceURI);
            }

            Assert.NotNull(childElement);
            Assert.Equal("ChildValue", childElement.InnerText);
            string childId = childElement.GetAttribute("id");
            Assert.Equal("123", childId);
        }

        [Fact (DisplayName = "FindElementByName should return an IEnumerable with one element, if there is only one matching child")]
        public void FindElementByNameTest()
        {
            XmlElement root = XmlElement.CreateElement("root");
            root.AddChildElementWithValue("node", "test1");
            IEnumerable<XmlElement> givenResult = root.FindChildElementsByName("node");
            Assert.Single(givenResult);
            Assert.Equal("test1", givenResult.First().InnerValue);
        }

        [Fact(DisplayName = "FindElementByName should return an IEnumerable with multiple element, if there more than one matching child")]
        public void FindElementByNameTest2()
        {
            XmlElement root = XmlElement.CreateElement("root");
            root.AddChildElementWithValue("node", "test1");
            root.AddChildElementWithValue("node", "test2");
            IEnumerable<XmlElement> givenResult = root.FindChildElementsByName("node");
            Assert.Equal(2, givenResult.Count());
            Assert.Single(givenResult.Where(n => n.InnerValue == "test1"));
            Assert.Single(givenResult.Where(n => n.InnerValue == "test2"));
        }

        [Fact(DisplayName = "FindElementByName should return an IEnumerable with multiple element, if there more than one matching child in a complex structure")]
        public void FindElementByNameTest3()
        {
            XmlElement root = XmlElement.CreateElement("root");
            XmlElement child1 = root.AddChildElement("subnode");
            child1.AddChildElementWithValue("node", "test1");
            child1.AddChildElementWithValue("node", "test2");
            XmlElement child2 = root.AddChildElement("subnode2");
            XmlElement child3 = child2.AddChildElement("subnode3");
            child3.AddChildElementWithValue("node", "test3", "pfx");
            IEnumerable<XmlElement> givenResult = root.FindChildElementsByName("node");
            Assert.Equal(3, givenResult.Count());
            Assert.Single(givenResult.Where(n => n.InnerValue == "test1"));
            Assert.Single(givenResult.Where(n => n.InnerValue == "test2"));
            Assert.Single(givenResult.Where(n => n.InnerValue == "test3"));
        }

        [Theory(DisplayName = "FindElementByName should return an empty IEnumerable, if there is no matching child")]
        [InlineData(null)]
        [InlineData("")]
        [InlineData(" ")]
        [InlineData("NODE")]
        [InlineData("node1")]
        [InlineData("test1")]
        public void FindElementByNameEmptyTest(string givenName)
        {
            XmlElement root = XmlElement.CreateElement("root");
            root.AddChildElementWithValue("node", "test1");
            IEnumerable<XmlElement> givenResult = root.FindChildElementsByName(givenName);
            Assert.Empty(givenResult);
        }

        [Fact(DisplayName = "FindElementByName should return an empty IEnumerable, if there are no child elements at all")]
        public void FindElementByNameEmptyTest2()
        {
            XmlElement root = XmlElement.CreateElement("root");
            IEnumerable<XmlElement> givenResult = root.FindChildElementsByName("node");
            Assert.Empty(givenResult);
        }

        [Fact(DisplayName = "FindElementByNameAndAttribute should return an IEnumerable with one element, if there is one matching child")]
        public void FindElementByNameAndAttributeTest()
        {
            XmlElement root = XmlElement.CreateElement("root");
            root.AddChildElementWithAttribute("node", "att1", "test1");
            IEnumerable<XmlElement> givenResult = root.FindChildElementsByNameAndAttribute("node", "att1", "test1");
            Assert.Single(givenResult);
            Assert.Equal("test1", givenResult.First().Attributes.First().Value);
        }

        [Fact(DisplayName = "FindElementByNameAndAttribute should return an IEnumerable with multiple elements, if there is more tahn one matching child")]
        public void FindElementByNameAndAttributeTest2()
        {
            XmlElement root = XmlElement.CreateElement("root");
            XmlElement child1 = root.AddChildElementWithAttribute("node", "att1", "test1");
            child1.InnerValue = "inner-value1";
            XmlElement child2 = root.AddChildElementWithAttribute("node", "att1", "no-match");
            child2.InnerValue = "inner-value2";
            XmlElement child3 = root.AddChildElementWithAttribute("node", "att1", "test1");
            child3.InnerValue = "inner-value3";
            IEnumerable<XmlElement> givenResult = root.FindChildElementsByNameAndAttribute("node", "att1", "test1");
            Assert.Equal(2, givenResult.Count());
            Assert.Single(givenResult.Where(n => n.InnerValue == "inner-value1"));
            Assert.Single(givenResult.Where(n => n.InnerValue == "inner-value3"));
        }

        [Fact(DisplayName = "FindElementByName should return an IEnumerable with multiple element, if there more than one matching child in a complex structure")]
        public void FindElementByNameAndAttributeTest3()
        {
            XmlElement root = XmlElement.CreateElement("root");
            XmlElement child1 = root.AddChildElement("subnode");
            XmlElement child1a = child1.AddChildElementWithValue("node", "test1");
            child1a.AddAttribute("att1", "test1");
            XmlElement child1b = child1.AddChildElementWithValue("node", "test2", "pfx");
            child1b.AddAttribute("att1", "test1");
            XmlElement child2 = root.AddChildElement("subnode2");
            child2.AddAttribute("node", "test1"); // should not match
            XmlElement child3 = child2.AddChildElement("subnode3");
            XmlElement child3a = child3.AddChildElementWithValue("node", "test3", "pfx");
            child3a.AddAttribute("att1", "test1");
            IEnumerable<XmlElement> givenResult = root.FindChildElementsByNameAndAttribute("node", "att1", "test1");
            Assert.Equal(3, givenResult.Count());
            Assert.Single(givenResult.Where(n => n.InnerValue == "test1"));
            Assert.Single(givenResult.Where(n => n.InnerValue == "test2"));
            Assert.Single(givenResult.Where(n => n.InnerValue == "test3"));
        }

        [Theory(DisplayName = "FindElementByNameAndAttribute should return an empty IEnumerable, if there is no matching child")]
        [InlineData(null, "att1", "test1")]
        [InlineData("", "att1", "test1")]
        [InlineData(" ", "att1", "test1")]
        [InlineData("NODE", "att1", "test1")]
        [InlineData("node1", "att1", "test1")]
        [InlineData("test1", "att1", "test1")]
        [InlineData("node", "att2", "test1")]
        [InlineData("node", "ATT1", "test1")]
        [InlineData("node", "att1", null)]
        [InlineData("node", "att1", "")]
        [InlineData("node", "att1", " ")]
        [InlineData("node", "att1", "TEST1")]
        [InlineData("node", "att1", "test2")]
        public void FindElementByNameAndAttributeEmptyTest(string givenTagName, string givenAttributeName, string givenAttributevalue)
        {
            XmlElement root = XmlElement.CreateElement("root");
            XmlElement child1 = root.AddChildElementWithValue("node", "test1");
            child1.AddAttribute("att", "test1");
            IEnumerable<XmlElement> givenResult = root.FindChildElementsByNameAndAttribute(givenTagName, givenAttributeName, givenAttributevalue);
            Assert.Empty(givenResult);
        }

        [Fact(DisplayName = "FindElementByNameAndAttribute should return an empty IEnumerable, if there are no child elements at all")]
        public void FindElementByNameAndAttributeEmptyTest2()
        {
            XmlElement root = XmlElement.CreateElement("root");
            IEnumerable<XmlElement> givenResult = root.FindChildElementsByNameAndAttribute("node", "att1", "test1");
            Assert.Empty(givenResult);
        }

    }
}
