/*
* NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
* Copyright Raphael Stoeckli © 2021
* This library is licensed under the MIT License.
* You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
*/

using System;
using System.Linq;
using System.Xml;

namespace NanoXLSX.LowLevel
{
    /// <summary>
    /// Static class with common util methods, used during reading XLSX files
    /// </summary>
    public static class ReaderUtils
    {
        /// <summary>
        /// Gets the XML attribute of the passed XML node by its name
        /// </summary>
        /// <param name="targetName">Name of the target attribute</param>
        /// <param name="node">XML node that contains the attribute</param>
        /// <param name="fallbackValue">Optional fallback value if the attribute was not found. Default is null</param>
        /// <returns>Attribute value as string or default value if not found (can be null)</returns>
        public static string GetAttribute(string targetName, XmlNode node, string fallbackValue = null)
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
    }
}
