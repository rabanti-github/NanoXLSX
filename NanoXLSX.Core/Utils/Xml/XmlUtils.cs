/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2025
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System.Collections.Generic;
using System.Text;

namespace NanoXLSX.Utils.Xml
{
    /// <summary>
    /// Class providing static methods to manipulate XML during packing or unpacking
    /// </summary>
    /// \remark <remarks>Methods in this class should only be used by the library components and not called by user code</remarks>
    public static class XmlUtils
    {
        /// <summary>
        /// Method to sanitize XML string values between two XML tags or in an attribute. Not considered are '&lt;', '&gt;' and '&amp;' since these characters are automatically escaped on writing XML 
        /// </summary>
        /// <param name="input">Input string to process</param>
        /// <returns>Escaped string</returns>
        /// \remark <remarks>Note: The XML specs allow characters up to the character value of 0x10FFFF. However, the C# char range is only up to 0xFFFF. NanoXLSX will neglect all values above this level in the sanitizing check. Illegal characters like 0x1 will be replaced with a white space (0x20)</remarks>
        public static string SanitizeXmlValue(string input)
        {
            if (input == null) { return ""; }
            var len = input.Length;
            var illegalCharacters = new List<int>(len);
            int i;
            for (i = 0; i < len; i++)
            {
                if (char.IsSurrogate(input[i]))
                {
                    if (i + 1 < input.Length && char.IsSurrogatePair(input[i], input[i + 1]))
                    {
                        // Valid surrogate pair; append both characters as-is.
                        i++; // Skip the next character.
                        continue;
                    }
                    else
                    {
                        illegalCharacters.Add(i);
                        continue;
                    }
                }
                if (input[i] < 0x9 || input[i] > 0xA && input[i] < 0xD || input[i] > 0xD && input[i] < 0x20 || input[i] > 0xD7FF && input[i] < 0xE000 || input[i] > 0xFFFD)
                {
                    illegalCharacters.Add(i);
                    continue;
                } // Note: XML specs allow characters up to 0x10FFFF. However, the C# char range is only up to 0xFFFF; Higher values are neglected here 
            }
            if (illegalCharacters.Count == 0)
            {
                return input;
            }

            var sb = new StringBuilder(len);
            var lastIndex = 0;
            len = illegalCharacters.Count;
            for (i = 0; i < len; i++)
            {
                sb.Append(input.Substring(lastIndex, illegalCharacters[i] - lastIndex));
                sb.Append(' '); // Whitespace as fall back on illegal character
                lastIndex = illegalCharacters[i] + 1;
            }
            sb.Append(input.Substring(lastIndex));
            return sb.ToString();
        }
    }
}
