/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2024
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */


using System.Collections.Generic;
using System.Text;

namespace NanoXLSX.Shared.Utils
{

    /// <summary>
    /// Class providing static methods to manipulate XML during packing or unpacking
    /// </summary>
    /// \remark <remarks>Methods in this class should only be used by the library components and not called by user code</remarks>
    public static class XmlUtils
    {

        /// <summary>
        /// Method to escape XML characters between two XML tags
        /// </summary>
        /// <param name="input">Input string to process</param>
        /// <returns>Escaped string</returns>
        /// \remark <remarks>Note: The XML specs allow characters up to the character value of 0x10FFFF. However, the C# char range is only up to 0xFFFF. NanoXLSX will neglect all values above this level in the sanitizing check. Illegal characters like 0x1 will be replaced with a white space (0x20)</remarks>
        public static string EscapeXmlChars(string input)
        {
            if (input == null) { return ""; }
            int len = input.Length;
            List<int> illegalCharacters = new List<int>(len);
            List<byte> characterTypes = new List<byte>(len);
            int i;
            for (i = 0; i < len; i++)
            {
                if (char.IsSurrogate(input[i]) && i + 1 < len && char.IsSurrogatePair(input[i], input[i + 1]))
                {
                    illegalCharacters.Add(i);
                    characterTypes.Add(4);
                    illegalCharacters.Add(i + 1); // Skip next character
                    characterTypes.Add(5);
                    i++;
                    continue;
                }
                if ((input[i] < 0x9) || (input[i] > 0xA && input[i] < 0xD) || (input[i] > 0xD && input[i] < 0x20) || (input[i] > 0xD7FF && input[i] < 0xE000) || (input[i] > 0xFFFD))
                {
                    illegalCharacters.Add(i);
                    characterTypes.Add(0);
                    continue;
                } // Note: XML specs allow characters up to 0x10FFFF. However, the C# char range is only up to 0xFFFF; Higher values are neglected here 
                if (input[i] == 0x3C) // <
                {
                    illegalCharacters.Add(i);
                    characterTypes.Add(1);
                }
                else if (input[i] == 0x3E) // >
                {
                    illegalCharacters.Add(i);
                    characterTypes.Add(2);
                }
                else if (input[i] == 0x26) // &
                {
                    illegalCharacters.Add(i);
                    characterTypes.Add(3);
                }
            }
            if (illegalCharacters.Count == 0)
            {
                return input;
            }

            StringBuilder sb = new StringBuilder(len);
            int lastIndex = 0;
            len = illegalCharacters.Count;
            for (i = 0; i < len; i++)
            {
                sb.Append(input.Substring(lastIndex, illegalCharacters[i] - lastIndex));
                if (characterTypes[i] == 0)
                {
                    sb.Append(' '); // Whitespace as fall back on illegal character
                }
                else if (characterTypes[i] == 1) // replace <
                {
                    sb.Append("&lt;");
                }
                else if (characterTypes[i] == 2) // replace >
                {
                    sb.Append("&gt;");
                }
                else if (characterTypes[i] == 3) // replace &
                {
                    sb.Append("&amp;");
                }
                else if (characterTypes[i] == 4)
                {
                    int codePoint = char.ConvertToUtf32(input[illegalCharacters[i]], input[illegalCharacters[i + 1]]);
                    sb.Append($"&#x{codePoint:X};");
                }
                else if (characterTypes[i] == 5)
                {
                    // Ignore character
                }
                lastIndex = illegalCharacters[i] + 1;
            }
            sb.Append(input.Substring(lastIndex));
            return sb.ToString();
        }

        /// <summary>
        /// Method to escape XML characters in an XML attribute
        /// </summary>
        /// <param name="input">Input string to process</param>
        /// <returns>Escaped string</returns>
        public static string EscapeXmlAttributeChars(string input)
        {
            input = EscapeXmlChars(input); // Sanitize string from illegal characters beside quotes
            input = input.Replace("\"", "&quot;");
            return input;
        }
    }
}
