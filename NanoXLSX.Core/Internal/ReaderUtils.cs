/*
* NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
* Copyright Raphael Stoeckli © 2024
* This library is licensed under the MIT License.
* You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
*/

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Linq.Expressions;
using System.Xml;

namespace NanoXLSX.Internal
{
    /// <summary>
    /// Static class with common util methods, used during reading XLSX files
    /// </summary>
    public static class ReaderUtils
    {
            /// <summary>
            /// Gets the XML attribute of the passed XML node by its name
            /// </summary>
            /// <param name="node">XML node that contains the attribute</param>
            /// <param name="targetName">Name of the target attribute</param>
            /// <param name="fallbackValue">Optional fallback value if the attribute was not found. Default is null</param>
            /// <returns>Attribute value as string or default value if not found (can be null)</returns>
            public static string GetAttribute(XmlNode node, string targetName, string fallbackValue = null)
        {
            if (node.Attributes == null || node.Attributes.Count == 0)
            {
                return fallbackValue;
            }

            foreach (XmlAttribute attribute in node.Attributes)
            {
                if (attribute.Name == targetName)
                {
                    return attribute.Value;
                }
            }

            return fallbackValue;
        }

        /// <summary>
        /// Gets the XML attribute from a child node of the passed XML node by its name and the name of the child node.
        /// This method simplifies the process of gathering one single child node attribute
        /// </summary>
        /// <param name="node">XML node that contains the child node</param>
        /// <param name="childNodeName">Name of the child node</param>
        /// <param name="attributeName">Name of the attribute in the child node</param>
        /// <param name="output">Value of the attribute as string or null if not found</param>
        /// <returns>True if found, otherwise false</returns>
        public static bool GetAttributeOfChild(XmlNode node, string childNodeName, string attributeName, out string output)
        {
            XmlNode childNode = GetChildNode(node, childNodeName);
            if (childNode != null)
            {
                output = GetAttribute(childNode, attributeName);
                return true;
            }
            output = null;
            return false;
        }

        /// <summary>
        /// Gets the specified child node
        /// </summary>
        /// <param name="node">XML node that contains child node</param>
        /// <param name="name">Name of the child node</param>
        /// <returns>Child node or null if not found</returns>
        public static XmlNode GetChildNode(XmlNode node, string name)
        {
            if (node != null && node.HasChildNodes)
            {
                return node.ChildNodes.Cast<XmlNode>().FirstOrDefault(c => c.LocalName.Equals(name, StringComparison.InvariantCultureIgnoreCase));
            }
            return null;
        }

        /// <summary>
        /// Checks whether the given node has the specified name
        /// </summary>
        /// <param name="node">XML node to check</param>
        /// <param name="name">Name to check</param>
        /// <returns>True if applying</returns>
        internal static bool IsNode(XmlNode node, string name)
        {
            return node.LocalName.Equals(name, StringComparison.InvariantCultureIgnoreCase);
        }

        /// <summary>
        /// Gets the prefix of an XML element with a target element name. If no prefix is defined, an empty string is returned
        /// </summary>
        /// <param name="document">XML document</param>
        /// <param name="targetName">Name of the target XML element</param>
        /// <returns>Prefix or empty</returns>
        internal static string DiscoverPrefix(XmlDocument document, string targetName)
        {
            foreach(XmlNode node in document.ChildNodes)
            {
                if (node.LocalName == targetName)
                {
                    return node.Prefix;
                }
            }
            return "";
        }

        /// <summary>
        /// Gets an XmlNodeList of the given name from the given document, either with or without prefix
        /// </summary>
        /// <param name="document">XML document</param>
        /// <param name="tagName">Name of the target XML element without prefix</param>
        /// <param name="prefix">prefix (without training colon) or empty</param>
        /// <returns>XmlNodeList of the found elements with the tag name</returns>
        internal static XmlNodeList GetElementsByTagName(XmlDocument document, string tagName, string prefix)
        {
            if (string.IsNullOrEmpty(prefix))
            {
                return document.GetElementsByTagName(tagName);
            }
            else
            {
                return document.GetElementsByTagName(prefix + ":" + tagName);
            }
        }

    }
}
