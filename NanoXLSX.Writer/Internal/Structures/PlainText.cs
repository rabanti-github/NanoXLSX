/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2025
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
using System;
using NanoXLSX.Interfaces;
using NanoXLSX.Utils;
using NanoXLSX.Utils.Xml;

namespace NanoXLSX.Internal.Structures
{
    /// <summary>
    /// Class to wrap unformatted strings as formattable text for the shared string table 
    /// </summary>
    /// \remark <remarks>This class is only for internal use. Use the high level API (e.g. class Workbook) to manipulate data and create Excel files</remarks>
    public class PlainText : IFormattableText
    {
        private const string TAG_NAME = "t";
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
            if (string.IsNullOrEmpty(Value))
            {
                return XmlElement.CreateElement(TAG_NAME);
            }
            string value = XmlUtils.SanitizeXmlValue(Value);
            value = ParserUtils.NormalizeNewLines(value);
            XmlElement element = null;
            if (Char.IsWhiteSpace(value, 0) || Char.IsWhiteSpace(value, value.Length - 1))
            {
                element = XmlElement.CreateElementWithAttribute(TAG_NAME, PRESERVE_ATTRIBUTE_NAME, PRESERVE_ATTRIBUTE_VALUE, "", PRESERVE_ATTRIBUTE_PREFIX_NAME);
            }
            else
            {
                element = XmlElement.CreateElement(TAG_NAME);
            }
            element.InnerValue = value;
            return element;
        }

        /// <summary>
        /// Constructor with value assignment
        /// </summary>
        /// <param name="value">Value to assign</param>
        public PlainText(string value)
        {
            this.Value = value;
        }


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
