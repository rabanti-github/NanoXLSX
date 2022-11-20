/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2022
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
using System;
using System.Text;
using NanoXLSX.Internal.Writers;
using NanoXLSX.Shared.Interfaces;
using NanoXLSX.Shared.Utils;

namespace NanoXLSX.Internal.Structures
{
    /// <summary>
    /// Class to wrap unformatted strings as formattable text for the stared string table 
    /// </summary>
    /// <remarks>This class is only used internally</remarks>
    public class PlainText : IFormattableText
    {
        private const string EMPTY_STRING = "<t></t>";
        private const string START_TAG = "<t>";
        private const string END_TAG = "</t>";
        private const string PRESERVE_START_TAG = "<t xml:space=\"preserve\">";

        public string Value { private get; set; }

        public void AddFormattedValue(StringBuilder sb)
        {
            if (string.IsNullOrEmpty(Value))
            {
                sb.Append(EMPTY_STRING);
            }
            else
            {
                string value = XmlUtils.EscapeXmlChars(Value);
                if (Char.IsWhiteSpace(value, 0) || Char.IsWhiteSpace(value, value.Length - 1))
                {
                    sb.Append(PRESERVE_START_TAG);
                }
                else
                {
                    sb.Append(START_TAG);
                }
                sb.Append(XlsxWriter.NormalizeNewLines(value)).Append(END_TAG);
            }
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
            return this.Value.Equals(((PlainText)obj).Value);
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
