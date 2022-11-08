/*
* NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
* Copyright Raphael Stoeckli © 2022
* This library is licensed under the MIT License.
* You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
*/

using System;
using System.Globalization;
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
        /// Parses a float independent of the culture info of the host
        /// </summary>
        /// <param name="rawValue">Raw number as string</param>
        /// <returns>Parsed float</returns>
        internal static float ParseFloat(String rawValue)
        {
            return float.Parse(rawValue, CultureInfo.InvariantCulture);
        }

        /// <summary>
        /// Parses an int independent of the culture info of the host
        /// </summary>
        /// <param name="rawValue">Raw number as string</param>
        /// <returns>Parsed int</returns>
        internal static int ParseInt(String rawValue)
        {
            return int.Parse(rawValue, CultureInfo.InvariantCulture);
        }

        /// <summary>
        /// Parses a bool as a binary number either based on an int (0/1) or a string expression (true/ false), independent of the culture info of the host
        /// </summary>
        /// <param name="rawValue">Raw number or expression as string</param>
        /// <returns>Parsed bool as number (0 = false, 1 = true)</returns>
        internal static int ParseBinaryBool(String rawValue)
        {
            if (string.IsNullOrEmpty(rawValue))
            {
                return 0;
            }
            int value;
            if (TryParseInt(rawValue, out value))
            {
                if (value >= 1)
                {
                    return 1;
                }
                else
                {
                    return 0;
                }
            }
            rawValue = rawValue.ToLower();
            if (rawValue == "true")
            {
                return 1;
            }
            else
            {
                return 0;
            }
        }

        /// <summary>
        /// Tries to parse an int independent of the culture info of the host
        /// </summary>
        /// <param name="rawvalue">Raw number as string</param>
        /// <param name="parsedValue">Parsed int</param>
        /// <returns>True, if the parsing was successful</returns>
        internal static bool TryParseInt(String rawvalue, out int parsedValue)
        {
            return int.TryParse(rawvalue, NumberStyles.Integer, CultureInfo.InvariantCulture, out parsedValue);
        }

    }
}
