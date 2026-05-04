/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System;
using System.Xml;

namespace NanoXLSX.Utils.Xml
{
    /// <summary>
    /// Static utility class providing helpers for forward-only XmlReader-based (SAX-style) parsing,
    /// used by all reader plug-ins to avoid XmlDocument DOM allocations.
    /// </summary>
    public static class XmlStreamUtils
    {
        /// <summary>
        /// Creates standardized XmlReaderSettings for all reader plug-ins:
        /// XmlResolver=null (security), whitespace/comment/PI nodes suppressed.
        /// </summary>
        public static XmlReaderSettings CreateSettings()
        {
            return new XmlReaderSettings
            {
                XmlResolver = null,
                IgnoreWhitespace = true,
                IgnoreComments = true,
                IgnoreProcessingInstructions = true,
            };
        }

        /// <summary>
        /// Returns true when the reader is positioned on a start element whose LocalName matches
        /// <paramref name="localName"/> (OrdinalIgnoreCase).
        /// </summary>
        public static bool IsElement(XmlReader reader, string localName)
        {
            return reader.NodeType == XmlNodeType.Element
                && reader.LocalName.Equals(localName, StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>
        /// Reads the string content of a simple leaf element (one that contains only text, no child
        /// elements). Leaves the reader positioned after the end element. Returns an empty string if
        /// the element is empty.
        /// </summary>
        public static string ReadElementText(XmlReader reader)
        {
            if (reader.IsEmptyElement)
            {
                return string.Empty;
            }
            return reader.ReadElementContentAsString();
        }
    }
}
