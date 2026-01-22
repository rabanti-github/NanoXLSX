/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using NanoXLSX.Interfaces;
using NanoXLSX.Utils;
using NanoXLSX.Utils.Xml;
using System;

namespace NanoXLSX.Internal.Structures
{
    /// <summary>
    /// Class to wrap unformatted strings as formattable text for the shared string table 
    /// </summary>
    /// \remark <remarks>This class is only for internal use. Use the high level API (e.g. class Workbook) to manipulate data and create Excel files</remarks>
    internal class PlainText : IFormattableText
    {
        private const string ITEM_TAG_NAME = "si";
        private const string TEXT_TAG_NAME = "t";
        private const string PRESERVE_ATTRIBUTE_NAME = "space";
        private const string PRESERVE_ATTRIBUTE_PREFIX_NAME = "xml";
        private const string PRESERVE_ATTRIBUTE_VALUE = "preserve";

        /// <summary>
        /// Unformatted Value (plain text)
        /// </summary>
        public string Value { private get; set; }

        /// <summary>
        /// Get the XmlElement (interface implementation)
        /// </summary>
        /// <returns>XmlElement instance</returns>
        public XmlElement GetXmlElement()
        {
            XmlElement siElement = XmlElement.CreateElement(ITEM_TAG_NAME);
            if (string.IsNullOrEmpty(Value))
            {
                siElement.AddChildElement(TEXT_TAG_NAME);
                return siElement;
            }
            string value = XmlUtils.SanitizeXmlValue(Value);
            value = ParserUtils.NormalizeNewLines(value);
            XmlElement element = null;
            if (Char.IsWhiteSpace(value, 0) || Char.IsWhiteSpace(value, value.Length - 1))
            {
                element = XmlElement.CreateElementWithAttribute(TEXT_TAG_NAME, PRESERVE_ATTRIBUTE_NAME, PRESERVE_ATTRIBUTE_VALUE, "", PRESERVE_ATTRIBUTE_PREFIX_NAME);
            }
            else
            {
                element = XmlElement.CreateElement(TEXT_TAG_NAME);
            }
            element.InnerValue = value;
            siElement.AddChildElement(element);
            return siElement;
        }

        /// <summary>
        /// Constructor with value assignment
        /// </summary>
        /// <param name="value">Value to assign</param>
        public PlainText(string value)
        {
            this.Value = value;
        }

        /// <summary>
        /// Determines whether the specified object is equal to the current object
        /// </summary>
        /// <param name="obj">Other object to compare</param>
        /// <returns>True if both objects are equal</returns>
        public override bool Equals(object obj)
        {
            if (this.Value == null && obj == null || (this.Value == null && ((PlainText)obj).Value == null))
            {
                return true;
            }
            else if (this.Value != null && !(obj is PlainText) || this.Value == null && ((PlainText)obj).Value != null)
            {
                return false;
            }
            return this.Value.Equals(((PlainText)obj).Value, StringComparison.Ordinal);
        }

        /// <summary>
        /// Gets the hash code based on the value
        /// </summary>
        /// <returns>Hash code</returns>
        public override int GetHashCode()
        {
            if (this.Value == null)
            {
                return 0;
            }
            return this.Value.GetHashCode();
        }

    }
}
