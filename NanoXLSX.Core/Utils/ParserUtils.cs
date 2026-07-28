/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System;
using System.Globalization;
using System.Linq;
using System.Text;

namespace NanoXLSX.Utils
{
    /// <summary>
    /// Class providing static methods to parse string values to specific types or to print object as language neutral string
    /// </summary>
    /// \remark <remarks>Methods in this class should only be used by the library components and not called by user code</remarks>
    public static class ParserUtils
    {
        #region constants

        /// <summary>
        /// Numeric format for ToString conversions. This format ensures that a numeric value is printed in a language neutral way.
        /// </summary>
        public const string NumericFormat = "G";

        /// <summary>
        /// Constant for number conversion. The invariant culture (represents mostly the US numbering scheme) ensures that no culture-specific 
        /// punctuations are used when converting numbers to strings, This is especially important for OOXML number values.
        /// See also: <a href="https://docs.microsoft.com/en-us/dotnet/api/system.globalization.cultureinfo.invariantculture?view=net-5.0">
        /// https://docs.microsoft.com/en-us/dotnet/api/system.globalization.cultureinfo.invariantculture?view=net-5.0</a>
        /// </summary>
        public static readonly CultureInfo InvariantCulture = CultureInfo.InvariantCulture;

        #endregion

        /// <summary>
        /// Determines whether a string starts with a specific value
        /// </summary>
        /// <param name="input">String to check</param>
        /// <param name="value">Value to be checked, whether it occurs at the beginning of the input string</param>
        /// <returns>True if the input string starts with the specified value</returns>
        public static bool StartsWith(string input, string value)
        {
            if (input == null && value == null)
            {
                return true;
            }
            else if (input == null && value != null)
            {
                return false;
            }
            else if (value == null)
            {
                return false;
            }
            return input.StartsWith(value, StringComparison.InvariantCulture);
        }

        /// <summary>
        /// Determines whether a string does not start with a specific value
        /// </summary>
        /// <param name="input">String to check</param>
        /// <param name="value">Value to be checked, whether it occurs not at the beginning of the input string</param>
        /// <returns>True if the input string does not starts with the specified value</returns>
        public static bool NotStartsWith(string input, string value)
        {
            return !StartsWith(input, value);
        }

        /// <summary>
        /// Transforms a string to upper case with null check and invariant culture
        /// </summary>
        /// <param name="input">String to transform</param>
        /// <returns>Upper case string</returns>
        public static string ToUpper(string input)
        {
            return !string.IsNullOrEmpty(input) ? input.ToUpper(InvariantCulture) : input;
        }

        /// <summary>
        /// Transforms a string to lower case with null check and invariant culture
        /// </summary>
        /// <param name="input">String to transform</param>
        /// <returns>Lower case string</returns>
        public static string ToLower(string input)
        {
            return !string.IsNullOrEmpty(input) ? input.ToLower(InvariantCulture) : input;
        }

        /// <summary>
        /// Transforms an integer to an invariant sting
        /// </summary>
        /// <param name="input">Integer to transform</param>
        /// <returns>Integer as string</returns>
        public static string ToString(int input)
        {
            return input.ToString(NumericFormat, InvariantCulture);
        }

        /// <summary>
        /// Transforms a float to an invariant sting
        /// </summary>
        /// <param name="input">Float to transform</param>
        /// <returns>Float as string</returns>
        public static string ToString(float input)
        {
            return input.ToString(NumericFormat, InvariantCulture);
        }

        /// <summary>
        /// Transforms a byte to an invariant sting
        /// </summary>
        /// <param name="input">Byte to transform</param>
        /// <returns>Byte as string</returns>
        public static string ToString(byte input)
        {
            return input.ToString(NumericFormat, InvariantCulture);
        }

        /// <summary>
        /// Transforms a sbyte to an invariant sting
        /// </summary>
        /// <param name="input">Sbyte to transform</param>
        /// <returns>Byte as string</returns>
        public static string ToString(sbyte input)
        {
            return input.ToString(NumericFormat, InvariantCulture);
        }

        /// <summary>
        /// Transforms a double to an invariant sting
        /// </summary>
        /// <param name="input">Double to transform</param>
        /// <returns>Double as string</returns>
        public static string ToString(double input)
        {
            return input.ToString(NumericFormat, InvariantCulture);
        }

        /// <summary>
        /// Transforms a decimal to an invariant sting
        /// </summary>
        /// <param name="input">Decimal to transform</param>
        /// <returns>Decimal as string</returns>
        public static string ToString(decimal input)
        {
            return input.ToString(NumericFormat, InvariantCulture);
        }

        /// <summary>
        /// Transforms a uint to an invariant sting
        /// </summary>
        /// <param name="input">Uint to transform</param>
        /// <returns>Uint as string</returns>
        public static string ToString(uint input)
        {
            return input.ToString(NumericFormat, InvariantCulture);
        }

        /// <summary>
        /// Transforms a long to an invariant sting
        /// </summary>
        /// <param name="input">Long to transform</param>
        /// <returns>Long as string</returns>
        public static string ToString(long input)
        {
            return input.ToString(NumericFormat, InvariantCulture);
        }

        /// <summary>
        /// Transforms a ulong to an invariant sting
        /// </summary>
        /// <param name="input">Ulong to transform</param>
        /// <returns>Ulong as string</returns>
        public static string ToString(ulong input)
        {
            return input.ToString(NumericFormat, InvariantCulture);
        }

        /// <summary>
        /// Transforms a short to an invariant sting
        /// </summary>
        /// <param name="input">Short to transform</param>
        /// <returns>Short as string</returns>
        public static string ToString(short input)
        {
            return input.ToString(NumericFormat, InvariantCulture);
        }

        /// <summary>
        /// Transforms a ushort to an invariant sting
        /// </summary>
        /// <param name="input">Ushort to transform</param>
        /// <returns>Ushort as string</returns>
        public static string ToString(ushort input)
        {
            return input.ToString(NumericFormat, InvariantCulture);
        }

        /// <summary>
        /// Transforms a given object to a string displayed as cached Values. The common known compatible numeric types, like int, float, sbyte etc. will be transformed to their appropriate string representations.  
        /// Bool will be either 0 or 1, or to TRUE or FALSE if convertBoolToNumber is set to false. 
        /// Date or TimeSpan will be transformed to a OADate (numeric) value.
        /// Null or empty will be transformed to 0. 
        /// If an unknown object type is passed, its own ToString() method will be used. 
        /// </summary>
        /// <param name="input">Object to transform</param>
        /// <param name="convertBoolToNumber">If set to true, a bool value will be TRUE or FALSE, otherwise 1 and 0. Default is true</param>
        /// <returns>Most appropriate OOXML string form given </returns>
        /// \remark <remarks>This method transforms values to the Excel-internal OOXML format. It is not meant as a generic ToString() method. Also do not pass nested objects like <see cref="Cell"/> or <see cref="FormulaData"/>, since they will be handled as unknown object types.</remarks>
        public static string ToCachedValueString(object input, bool convertBoolToNumber = true)
        {
            if (input == null) { return string.Empty; }
            else if (input is string)
            {
                return input as string;
            }
            else if (input is bool)
            {
                if (convertBoolToNumber)
                {
                    return (bool)input ? "1" : "0";
                }
                else
                {
                    return (bool)input ? "TRUE" : "FALSE";
                }
            }
            else if (input is bool) { return ToString((byte)input); }
            else if (input is sbyte) { return ToString((sbyte)input); }
            else if (input is decimal) { return ToString((decimal)input); }
            else if (input is double) { return ToString((double)input); }
            else if (input is int) { return ToString((int)input); }
            else if (input is uint) { return ToString((uint)input); }
            else if (input is long) { return ToString((ulong)input); }
            else if (input is ulong) { return ToString((ulong)input); }
            else if (input is short) { return ToString((ushort)input); }
            else if (input is ushort) { return ToString((ushort)input); }
            else if (input is DateTime)
            {
                return DataUtils.GetOADateTimeString((DateTime)input);
            }
            else if (input is TimeSpan)
            {
                return DataUtils.GetOATimeString((TimeSpan)input);
            }
            else
            {
                return input.ToString(); // Generic string
            }
        }


        /// <summary>
        /// Normalizes all newlines of a string to CR+LF
        /// </summary>
        /// <param name="value">Input value</param>
        /// <returns>Normalized value</returns>
        public static string NormalizeNewLines(string value)
        {
            if (value == null || (!value.Contains('\n') && !value.Contains('\r')))
            {
                return value;
            }
            return value.Replace("\n\r", "\n").Replace("\r\n", "\n").Replace("\r", "\n").Replace("\n", "\r\n");
        }

        /// <summary>
        /// Parses a float independent from the culture info of the host
        /// </summary>
        /// <param name="rawValue">Raw number as string</param>
        /// <returns>Parsed float</returns>
        /// \remark <remarks>The method does not check the validity and will cause an error if an invalid value is passed</remarks>
        public static float ParseFloat(string rawValue)
        {
            return float.Parse(rawValue, InvariantCulture);
        }

        /// <summary>
        /// Parses an int independent from the culture info of the host
        /// </summary>
        /// <param name="rawValue">Raw number as string</param>
        /// <returns>Parsed int</returns>
        /// \remark <remarks>The method does not check the validity and will cause an error if an invalid value is passed</remarks>
        public static int ParseInt(string rawValue)
        {
            return int.Parse(rawValue, NumberStyles.Any, InvariantCulture);
        }

        /// <summary>
        /// Parses a double independent from the culture info of the host
        /// </summary>
        /// <param name="rawValue">Raw number as string</param>
        /// <returns>Parsed int</returns>
        /// \remark <remarks>The method does not check the validity and will cause an error if an invalid value is passed</remarks>
        public static double ParseDouble(string rawValue)
        {
            return double.Parse(rawValue, InvariantCulture);
        }

        /// <summary>
        /// Parses a bool as a binary number either based on an int (0/1) or a string expression (true/ false), independent of the culture info of the host
        /// </summary>
        /// <param name="rawValue">Raw number or expression as string</param>
        /// <returns>Parsed bool as number (0 = false, 1 = true)</returns>
        public static int ParseBinaryBool(String rawValue)
        {
            if (string.IsNullOrEmpty(rawValue))
            {
                return 0;
            }
            int value;
            if (TryParseInt(rawValue, out value))
            {
                return value >= 1 ? 1 : 0;
            }
            bool regularBool;
            if (TryParseBool(rawValue, out regularBool))
            {
                return regularBool ? 1 : 0;
            }
            return 0;
        }

        /// <summary>
        /// Tries to parse a bool from its string name (true/false), independent from the case
        /// </summary>
        /// <param name="rawValue">Raw bool as string</param>
        /// <param name="parsedValue">Parsed bool</param>
        /// <returns>True, if the parsing was successful</returns>
        /// /remark <remarks>Integer values, like 0 or 1 will return false</remarks>
        public static bool TryParseBool(string rawValue, out bool parsedValue)
        {
            return bool.TryParse(rawValue, out parsedValue);
        }

        /// <summary>
        /// Tries to parse an int independent of the culture info of the host
        /// </summary>
        /// <param name="rawValue">Raw number as string</param>
        /// <param name="parsedValue">Parsed int</param>
        /// <returns>True, if the parsing was successful</returns>
        public static bool TryParseInt(string rawValue, out int parsedValue)
        {
            return int.TryParse(rawValue, NumberStyles.Integer, InvariantCulture, out parsedValue);
        }

        /// <summary>
        /// Tries to parse an unsigned int (uint) independent from the culture info of the host
        /// </summary>
        /// <param name="rawValue">Raw number as string</param>
        /// <param name="parsedValue">Parsed uint</param>
        /// <returns>True, if the parsing was successful</returns>
        public static bool TryParseUint(string rawValue, out uint parsedValue)
        {
            return uint.TryParse(rawValue, NumberStyles.Integer, InvariantCulture, out parsedValue);
        }

        /// <summary>
        /// Tries to parse a long independent from the culture info of the host
        /// </summary>
        /// <param name="rawValue">Raw number as string</param>
        /// <param name="parsedValue">Parsed long</param>
        /// <returns>True, if the parsing was successful</returns>
        public static bool TryParseLong(string rawValue, out long parsedValue)
        {
            return long.TryParse(rawValue, NumberStyles.Integer, InvariantCulture, out parsedValue);
        }

        /// <summary>
        /// Tries to parse an unsigned long (ulong) independent from the culture info of the host
        /// </summary>
        /// <param name="rawValue">Raw number as string</param>
        /// <param name="parsedValue">Parsed ulong</param>
        /// <returns>True, if the parsing was successful</returns>
        public static bool TryParseUlong(string rawValue, out ulong parsedValue)
        {
            return ulong.TryParse(rawValue, NumberStyles.Integer, InvariantCulture, out parsedValue);
        }

        /// <summary>
        /// Tries to parse a float (with any parsing style) independent from the culture info of the host
        /// </summary>
        /// <param name="rawValue">Raw number as string</param>
        /// <param name="parsedValue">Parsed float</param>
        /// <returns>True, if the parsing was successful</returns>
        public static bool TryParseFloat(string rawValue, out float parsedValue)
        {
            return float.TryParse(rawValue, NumberStyles.Any, CultureInfo.InvariantCulture, out parsedValue);
        }

        /// <summary>
        /// Tries to parse a decimal (with float parsing style) independent from the culture info of the host
        /// </summary>
        /// <param name="rawValue">Raw number as string</param>
        /// <param name="parsedValue">Parsed decimal</param>
        /// <returns>True, if the parsing was successful</returns>
        public static bool TryParseDecimal(string rawValue, out decimal parsedValue)
        {
            return decimal.TryParse(rawValue, NumberStyles.Float, InvariantCulture, out parsedValue);
        }

        /// <summary>
        /// Tries to parse a double (with any parsing style) independent from the culture info of the host
        /// </summary>
        /// <param name="rawValue">Raw number as string</param>
        /// <param name="parsedValue">Parsed double</param>
        /// <returns>True, if the parsing was successful</returns>
        public static bool TryParseDouble(string rawValue, out double parsedValue)
        {
            return double.TryParse(rawValue, NumberStyles.Any, InvariantCulture, out parsedValue);
        }

        /// <summary>
        /// Tries to parse a raw string as an Excel formula string constant.
        /// Escaped double quotes (<c>""</c>) are converted to single double quotes (<c>"</c>).
        /// </summary>
        /// <param name="expression">Raw string expression.</param>
        /// <param name="value">Parsed and unescaped string value as out parameter.</param>
        /// <param name="enclosingQuotesRemoved">If true, the enclosing leading and trailing double quotes have already been removed. Default is false.</param>
        /// <returns>True if the expression is a valid formula string constant; otherwise false.</returns>
        public static bool TryParseFormulaStringConstant(string expression, out string value, bool enclosingQuotesRemoved = false)
        {
            value = null;
            if (expression == null)
            {
                return false;
            }

            int startIndex;
            int endIndex;
            if (enclosingQuotesRemoved)
            {
                // The complete expression represents the content inside the string literal.
                startIndex = 0;
                endIndex = expression.Length;
            }
            else
            {
                // A complete formula string constant requires enclosing double quotes.
                if (expression.Length < 2
                    || expression[0] != '"'
                    || expression[expression.Length - 1] != '"')
                {
                    return false;
                }
                startIndex = 1;
                endIndex = expression.Length - 1;
            }

            StringBuilder builder = new StringBuilder(endIndex - startIndex);
            for (int i = startIndex; i < endIndex; i++)
            {
                char current = expression[i];
                if (current != '"')
                {
                    builder.Append(current);
                    continue;
                }
                // Quotes inside an Excel string constant must always occur as a pair.
                if (i + 1 >= endIndex || expression[i + 1] != '"')
                {
                    return false;
                }
                builder.Append('"');
                i++;
            }

            value = builder.ToString();
            return true;
        }

        /// <summary>
        /// Tries to parse a qualifies worksheet name and address or range expression from a raw string.
        /// </summary>
        /// <param name="expression">Raw expression to parse</param>
        /// <param name="worksheetName">Resolved worksheet name as out parameter</param>
        /// <param name="reference">Resolved address or range expression as string (out parameter)</param>
        /// <returns>True if the expression is a valid worksheet name with attached reference (address / range); otherwise false.</returns>
        /// \remark <remarks>This method cannot detect worksheet names and references within formulas or cell calculations. Such an expression has to be tokenized first.</remarks>
        public static bool TryParseWorksheetQualifiedReference(string expression,  out string worksheetName, out string reference)
        {
            worksheetName = null;
            reference = null;

            if (string.IsNullOrEmpty(expression))
            {
                return false;
            }

            if (expression[0] != '\'')
            {
                int separatorIndex = expression.IndexOf('!');
                if (separatorIndex <= 0 || separatorIndex == expression.Length - 1)
                {
                    return false;
                }

                worksheetName = expression.Substring(0, separatorIndex);
                reference = expression.Substring(separatorIndex + 1);
                return true;
            }

            StringBuilder builder = new StringBuilder();
            for (int i = 1; i < expression.Length; i++)
            {
                char current = expression[i];
                if (current != '\'')
                {
                    builder.Append(current);
                    continue;
                }

                // Two apostrophes inside a quoted worksheet name represent one literal apostrophe.
                if (i + 1 < expression.Length && expression[i + 1] == '\'')
                {
                    builder.Append('\'');
                    i++;
                    continue;
                }

                // A single apostrophe closes the quoted worksheet name. It must immediately be followed by the reference separator.
                if (i + 1 >= expression.Length || expression[i + 1] != '!')
                {
                    return false;
                }

                if (i + 2 >= expression.Length)
                {
                    return false;
                }

                worksheetName = builder.ToString();
                reference = expression.Substring(i + 2);
                return true;
            }
            // No closing apostrophe was found.
            return false;
        }

    }
}
