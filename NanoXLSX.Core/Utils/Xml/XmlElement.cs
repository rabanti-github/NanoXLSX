/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way  
 * Copyright Raphael Stoeckli © 2025
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System.Collections.Generic;
using System.Linq;
using System.Xml;

namespace NanoXLSX.Utils.Xml
{
    /// <summary>
    /// Class representing an internally used XML element / node
    /// </summary>
    public class XmlElement
    {
        private readonly bool hasPrefix = false;
        private bool hasNameSpaces = false;
        private bool hasDefaultNameSpace = false;
        private bool hasAttributes = false;
        private bool hasInnerValue = false;
        private bool hasChildren = false;
        private string innerValue;
        private string defaultXmlNsUri;

        /// <summary>
        /// Prefix of the element. If not defined, the prefix will be an empty string
        /// </summary>
        public string Prefix { get; set; }
        /// <summary>
        /// Name of the element (without prefix)
        /// </summary>
        public string Name { get; private set; }
        /// <summary>
        /// List of child elements. If none, the list is null
        /// </summary>
        public List<XmlElement> Children { get; private set; }
        /// <summary>
        /// List of attributes of this element. If none, the list is null
        /// </summary>
        public HashSet<XmlAttribute> Attributes { get; private set; }
        /// <summary>
        /// Map of prefixes and corresponding name space URIs of this element
        /// </summary>
        public Dictionary<string, string> PrefixNameSpaceMap { get; private set; }

        /// <summary>
        /// Gets or set the inner value of the element
        /// </summary>
        public string InnerValue
        {
            get => innerValue;
            set
            {
                if (string.IsNullOrEmpty(value))
                {
                    innerValue = null;
                    hasInnerValue = false;
                }
                else
                {
                    innerValue = value;
                    hasInnerValue = true;
                }
            }
        }

        /// <summary>
        /// Constructor with parameters
        /// </summary>
        /// <param name="name">Name of the element</param>
        /// <param name="prefix">Prefix of the element</param>
        internal XmlElement(string name, string prefix)
        {
            this.Name = name;
            this.Prefix = prefix;
            this.hasPrefix = !string.IsNullOrEmpty(prefix);
        }

        /// <summary>
        /// Method to add a name space as element attribute.
        /// Make sure not to add 'xmlns' as prefix since this is usually only the default name space and will be added implicitly when defined by <see cref="AddDefaultXmlNameSpace(string)"/>
        /// </summary>
        /// <param name="prefix">Prefix of the name space</param>
        /// <param name="rootNameSpace">Root name space (usually 'xmlns'). This value can also be empty</param>
        /// <param name="uri">URI of the name space</param>
        internal void AddNameSpaceAttribute(string prefix, string rootNameSpace, string uri)
        {
            if (string.IsNullOrEmpty(prefix) || string.IsNullOrEmpty(uri))
            {
                return;
            }
            if (PrefixNameSpaceMap == null)
            {
                PrefixNameSpaceMap = new Dictionary<string, string>();
            }
            if (!PrefixNameSpaceMap.ContainsKey(prefix))
            {
                PrefixNameSpaceMap.Add(prefix, uri);
            }
            hasNameSpaces = true;
            AddAttribute(prefix, uri, rootNameSpace);
        }

        /// <summary>
        /// Method to add the default name space  URI of the current element. 
        /// </summary>
        /// <param name="defaultXmlNsUri">URI to be defined as default name space</param>
        internal void AddDefaultXmlNameSpace(string defaultXmlNsUri)
        {
            this.defaultXmlNsUri = defaultXmlNsUri;
            hasDefaultNameSpace = true;
        }

        /// <summary>
        /// Method to add an attribute to the element
        /// </summary>
        /// <param name="name">Attribute name</param>
        /// <param name="value">Attribute value</param>
        /// <param name="prefix">Optional attribute prefix</param>
        internal void AddAttribute(string name, string value, string prefix = "")
        {
            if (!hasAttributes)
            {
                Attributes = new HashSet<XmlAttribute>();
                hasAttributes = true;
            }
            Attributes.Add(XmlAttribute.CreateAttribute(name, value, prefix));
        }

        /// <summary>
        /// Method to add an attribute to the element
        /// </summary>
        /// <param name="nullableAttribute">Nullable attribute instance. If not defined, nothing will be added</param>
        internal void AddAttribute(XmlAttribute? nullableAttribute)
        {
            if (!nullableAttribute.HasValue)
            {
                return;
            }
            if (!hasAttributes)
            {
                Attributes = new HashSet<XmlAttribute>();
                hasAttributes = true;
            }
            Attributes.Add(nullableAttribute.Value);
        }

        /// <summary>
        /// Method to add a enumeration of attributes to the element
        /// </summary>
        /// <param name="attributes">IEnumerable of Attributes to add. If null or empty, nothing will be added</param>
        internal void AddAttributes(IEnumerable<XmlAttribute> attributes)
        {
            if (attributes == null | !attributes.Any())
            {
                return;
            }
            if (!hasAttributes)
            {
                Attributes = new HashSet<XmlAttribute>();
                hasAttributes = true;
            }
            foreach (XmlAttribute attribute in attributes)
            {
                Attributes.Add(attribute);
            }
        }

        /// <summary>
        /// Method to add A child element with one attribute to the current element
        /// </summary>
        /// <param name="name">Name of the child element</param>
        /// <param name="attributeName">Attribute name, added to the child element</param>
        /// <param name="attributeValue">Attribute value, added to the child element</param>
        /// <param name="namePrefix">Optional prefix of the child element</param>
        /// <param name="attributePrefix">Optional prefix of the attribute, added to the child element</param>
        /// <returns>Instance of the added child element</returns>
        internal XmlElement AddChildElementWithAttribute(string name, string attributeName, string attributeValue, string namePrefix = "", string attributePrefix = "")
        {
            XmlElement childElement = CreateElementWithAttribute(name, attributeName, attributeValue, namePrefix, attributePrefix);
            AddChildElement(childElement);
            return childElement;
        }

        /// <summary>
        /// Method to add A child element with an inner value
        /// </summary>
        /// <param name="name">Name of the child element</param>
        /// <param name="innerValue">Inner (text) value of the child element</param>
        /// <param name="prefix">Optional prefix of the child element</param>
        /// <returns>Instance of the added child element</returns>
        internal XmlElement AddChildElementWithValue(string name, string innerValue, string prefix = "")
        {
            if (string.IsNullOrEmpty(innerValue))
            {
                return null; // Omit empty nodes
            }
            XmlElement childElement = CreateElement(name, prefix);
            childElement.InnerValue = innerValue;
            AddChildElement(childElement);
            return childElement;
        }

        /// <summary>
        /// Method to add a child element to the current one
        /// </summary>
        /// <param name="name">Name of the child element</param>
        /// <param name="prefix">Optional prefix of the child element</param>
        /// <returns>Instance of the added child element</returns>
        internal XmlElement AddChildElement(string name, string prefix = "")
        {
            XmlElement childElement = CreateElement(name, prefix);
            AddChildElement(childElement);
            return childElement;
        }

        /// <summary>
        /// Method to add a child element to the current one
        /// </summary>
        /// <param name="xmlElement">Nullable child element instance. If null, nothing will be added</param>
        internal void AddChildElement(XmlElement xmlElement)
        {
            if (xmlElement == null)
            {
                return;
            }
            if (!hasChildren)
            {
                Children = new List<XmlElement>();
                hasChildren = true;
            }
            Children.Add(xmlElement);
        }

        /// <summary>
        /// Method to add an enumeration of child element to the current one
        /// </summary>
        /// <param name="xmlElements">IEnumerable of child elements to be added. If null or empty, nothing will be added</param>
        internal void AddChildElements(IEnumerable<XmlElement> xmlElements)
        {
            if (xmlElements == null || !xmlElements.Any())
            {
                return;
            }
            if (!hasChildren)
            {
                Children = new List<XmlElement>();
                hasChildren = true;
            }
            Children.AddRange(xmlElements);
        }

        /// <summary>
        /// Method to create an XML element
        /// </summary>
        /// <param name="name">Name of the element</param>
        /// <param name="prefix">Optional prefix of the element</param>
        /// <returns>Element instance</returns>
        internal static XmlElement CreateElement(string name, string prefix = "")
        {
            return new XmlElement(name, prefix);
        }

        /// <summary>
        /// Method to create an XML element with one attribute
        /// </summary>
        /// <param name="name">Name of the element</param>
        /// <param name="attributeName">Attribute name</param>
        /// <param name="attributeValue">Attribute value</param>
        /// <param name="namePrefix">Optional prefix of the attribute</param>
        /// <param name="attributePrefix"></param>
        /// <returns>Element instance</returns>
        internal static XmlElement CreateElementWithAttribute(string name, string attributeName, string attributeValue, string namePrefix = "", string attributePrefix = "")
        {
            XmlElement element = new XmlElement(name, namePrefix);
            element.Attributes = new HashSet<XmlAttribute>();
            element.Attributes.Add(XmlAttribute.CreateAttribute(attributeName, attributeValue, attributePrefix));
            element.hasAttributes = true;
            return element;
        }

        /// <summary>
        /// Transforms this custom XmlElement (and its children) into a standard XmlDocument.
        /// </summary>
        /// <returns>A new XmlDocument representing the hierarchical XML structure.</returns>
        public XmlDocument TransformToDocument()
        {
            XmlDocument doc = new XmlDocument();
            XmlNamespaceManager nsManager = new XmlNamespaceManager(doc.NameTable);
            if (hasNameSpaces)
            {
                foreach (KeyValuePair<string, string> nameSpace in PrefixNameSpaceMap)
                {
                    if (nameSpace.Key == "xmlns")
                    {
                        continue;
                    }
                    nsManager.AddNamespace(nameSpace.Key, nameSpace.Value);
                }
            }
            // Create the root element from this instance recursively.
            System.Xml.XmlElement rootElement = null;
            if (hasDefaultNameSpace)
            {
                rootElement = CreateXmlElement(doc, this, nsManager, defaultXmlNsUri);
            }
            else
            {
                rootElement = CreateXmlElement(doc, this, nsManager);
            }
            doc.AppendChild(rootElement);


            return doc;
        }

        /// <summary>
        /// Recursively creates a System.Xml.XmlElement from the custom XmlElement.
        /// </summary>
        /// <param name="doc">The XmlDocument to which the element belongs.</param>
        /// <param name="customElement">The custom XmlElement to convert.</param>
        /// <param name="nsManager">XML name space manager instance</param>
        /// <param name="defaultXmlNsUri">Optional URI of the default XML namespace URI</param>
        /// <returns>A System.Xml.XmlElement representing the custom element.</returns>
        private System.Xml.XmlElement CreateXmlElement(XmlDocument doc, XmlElement customElement, XmlNamespaceManager nsManager, string defaultXmlNsUri = null)
        {
            System.Xml.XmlElement xmlElem;
            if (customElement.hasPrefix)
            {
                xmlElem = doc.CreateElement(customElement.Prefix, customElement.Name, nsManager.LookupNamespace(customElement.Prefix));
            }
            else
            {
                if (string.IsNullOrEmpty(defaultXmlNsUri))
                {
                    xmlElem = doc.CreateElement(customElement.Name);
                }
                else
                {
                    xmlElem = doc.CreateElement(customElement.Name, defaultXmlNsUri);
                }
            }

            // Add attributes
            if (customElement.hasAttributes)
            {
                foreach (var attr in customElement.Attributes)
                {
                    if (attr.HasPrefix)
                    {
                        System.Xml.XmlAttribute xmlAttr = doc.CreateAttribute(attr.Prefix, attr.Name, nsManager.LookupNamespace(attr.Prefix));
                        xmlAttr.Value = attr.Value;
                        xmlElem.Attributes.Append(xmlAttr);
                    }
                    else
                    {
                        xmlElem.SetAttribute(attr.Name, attr.Value);
                    }
                }
            }

            // Set inner text if available
            if (customElement.hasInnerValue)
            {
                xmlElem.InnerText = customElement.InnerValue;
            }

            // Process children recursively.
            if (customElement.hasChildren)
            {
                foreach (var child in customElement.Children)
                {
                    System.Xml.XmlElement childXmlElem = CreateXmlElement(doc, child, nsManager, defaultXmlNsUri);
                    xmlElem.AppendChild(childXmlElem);
                }
            }
            return xmlElem;
        }
    }
}
