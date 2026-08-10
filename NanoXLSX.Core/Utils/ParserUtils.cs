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

        #region primitiveParsing
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
        /// Tries to parse a double independent from the culture info of the host
        /// </summary>
        /// <param name="rawValue">Raw number as string</param>
        /// <param name="parsedValue">Parsed double</param>
        /// <param name="numberStyles">Permitted number styles. Default is <see cref="NumberStyles.Any"/></param>
        /// <returns>True, if the parsing was successful</returns>
        public static bool TryParseDouble(string rawValue, out double parsedValue, NumberStyles numberStyles = NumberStyles.Any)
        {
            return double.TryParse(rawValue, numberStyles, InvariantCulture, out parsedValue);
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
        /// Determines whether the passed character is a ASCII digit character (0-9)
        /// </summary>
        /// <param name="character">Character to check</param>
        /// <returns>True if an ASCII character, otherwise false</returns>
        public static bool IsAsciiDigit(char character)
        {
            return character >= '0' && character <= '9';
        }
        #endregion

        #region formulaParsing
        /// <summary>
        /// Transforms a given object to a string displayed as cached Values. The common known compatible numeric types, like int, float, sbyte etc. will be transformed to their appropriate string representations.  
        /// Bool will be either 0 or 1, or to TRUE or FALSE if convertBoolToNumber is set to false. 
        /// Date or TimeSpan will be transformed to a OADate (numeric) value.
        /// Null or an empty string will be transformed to 0.
        /// If an unknown object type is passed, its own ToString() method will be used. 
        /// </summary>
        /// <param name="input">Object to transform</param>
        /// <param name="convertBoolToNumber">If set to true, a bool value will be 1 or 0, otherwise TRUE or FALSE. Default is true</param>
        /// <returns>Most appropriate OOXML string form given </returns>
        /// <exception cref="Exceptions.FormatException">Thrown if an invalid Date value was passed. See method <see cref="DataUtils.GetOADateTimeString(DateTime)"/> for details</exception>
        /// \remark <remarks>This method transforms values to the Excel-internal OOXML format. It is not meant as a generic ToString() method. Also do not pass nested objects like <see cref="Cell"/> or <see cref="FormulaData"/>, since they will be handled as unknown object types.</remarks>
        public static string ToCachedValueString(object input, bool convertBoolToNumber = true)
        {
            if (input == null) { return "0"; }
            else if (input is string)
            {
                string stringValue = input as string;
                return string.IsNullOrEmpty(stringValue) ? "0" : stringValue;
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
            else if (input is byte) { return ToString((byte)input); }
            else if (input is sbyte) { return ToString((sbyte)input); }
            else if (input is decimal) { return ToString((decimal)input); }
            else if (input is double) { return ToString((double)input); }
            else if (input is float) { return ToString((float)input); }
            else if (input is int) { return ToString((int)input); }
            else if (input is uint) { return ToString((uint)input); }
            else if (input is long) { return ToString((long)input); }
            else if (input is ulong) { return ToString((ulong)input); }
            else if (input is short) { return ToString((short)input); }
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
        #endregion

        #region referenceParsing
        /// <summary>
        /// Tries to parse a qualifies worksheet name and address or range expression from a raw string.
        /// </summary>
        /// <param name="expression">Raw expression to parse</param>
        /// <param name="worksheetName">Resolved worksheet name as out parameter</param>
        /// <param name="reference">Resolved address or range expression as string (out parameter)</param>
        /// <returns>True if the expression is a valid worksheet name with attached reference (address / range); otherwise false.</returns>
        /// \remark <remarks>This method cannot detect worksheet names and references within formulas or cell calculations. Such an expression has to be tokenized first.</remarks>
        public static bool TryParseWorksheetQualifiedReference(string expression, out string worksheetName, out string reference)
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
        #endregion

        #region externalReferenceParsing

        /// <summary>
        /// Determines whether a formula contains an external workbook reference without parsing the formula.
        /// String literals, structured references, and relative R1C1 references are ignored.
        /// </summary>
        /// <param name="formulaExpression">Formula expression without a leading equal sign.</param>
        /// <returns>True if an external workbook reference was found.</returns>
        internal static bool ContainsExternalReference(string formulaExpression)
        {
            if (string.IsNullOrEmpty(formulaExpression) || formulaExpression.IndexOf('[') < 0)
            {
                return false;
            }

            bool inStringLiteral = false;
            for (int i = 0; i < formulaExpression.Length; i++)
            {
                char current = formulaExpression[i];
                if (current == '"')
                {
                    if (inStringLiteral && i + 1 < formulaExpression.Length && formulaExpression[i + 1] == '"')
                    {
                        i++;
                        continue;
                    }
                    inStringLiteral = !inStringLiteral;
                    continue;
                }
                if (inStringLiteral || current != '[')
                {
                    continue;
                }

                int closingBracket = formulaExpression.IndexOf(']', i + 1);
                if (closingBracket <= i + 1)
                {
                    continue;
                }

                bool hasWorksheetName = false;
                for (int j = closingBracket + 1; j < formulaExpression.Length; j++)
                {
                    char referenceCharacter = formulaExpression[j];
                    if (referenceCharacter == '!')
                    {
                        if (hasWorksheetName)
                        {
                            return true;
                        }
                        break;
                    }
                    if (referenceCharacter == '[' || referenceCharacter == ']' || referenceCharacter == '"'
                        || referenceCharacter == '+' || referenceCharacter == '-' || referenceCharacter == '*'
                        || referenceCharacter == '/' || referenceCharacter == '^' || referenceCharacter == '&'
                        || referenceCharacter == '=' || referenceCharacter == '<' || referenceCharacter == '>'
                        || referenceCharacter == ',' || referenceCharacter == ';' || referenceCharacter == '('
                        || referenceCharacter == ')' || referenceCharacter == '{' || referenceCharacter == '}')
                    {
                        break;
                    }
                    if (!char.IsWhiteSpace(referenceCharacter) && referenceCharacter != '\'')
                    {
                        hasWorksheetName = true;
                    }
                }
                i = closingBracket;
            }
            return false;
        }

        /// <summary>
        /// Determines whether a passed identifier is a valid external link identifier (internal representation, e.g."[2]") 
        /// </summary>
        /// <param name="identifier">Expression to check</param>
        /// <returns>True if a valid external link ID, otherwise false</returns>
        internal static bool IsValidExternalLinkId(string identifier)
        {
            if (string.IsNullOrEmpty(identifier) ||
                identifier.Length < 3 ||
                identifier[0] != '[' ||
                identifier[identifier.Length - 1] != ']')
            {
                return false;
            }

            for (int i = 1; i < identifier.Length - 1; i++)
            {
                if (!IsAsciiDigit(identifier[i]))
                {
                    return false;
                }
            }
            return true;
        }

        /// <summary>
        /// Tries to read an external workbook identifier (internal representation, e.g. "[1]") beginning at the specified position. 
        /// The out parameter is the length of the token.
        /// </summary>
        /// <param name="expression">Expression where the identifier is supposed to be</param>
        /// <param name="startIndex">Start index in the expression</param>
        /// <param name="identifierLength">Length of the identifier token as out parameter</param>
        /// <returns>True if the external link ID could be read, otherwise false</returns>
        /// \remark <remarks>The actual reading can be made, based on the output, with <see cref="string.Substring(int, int)"/></remarks>
        internal static bool TryReadExternalLinkId(string expression, int startIndex, out int identifierLength)
        {
            identifierLength = 0;
            if (startIndex < 0 || startIndex >= expression.Length || expression[startIndex] != '[')
            {
                return false;
            }

            int currentIndex = startIndex + 1;
            if (currentIndex >= expression.Length ||
                !IsAsciiDigit(expression[currentIndex]))
            {
                return false;
            }

            do
            {
                currentIndex++;
            }
            while (currentIndex < expression.Length && IsAsciiDigit(expression[currentIndex]));
            if (currentIndex >= expression.Length || expression[currentIndex] != ']')
            {
                return false;
            }

            int closingBracketIndex = currentIndex;
            if (!HasValidPrefixBoundary(expression, startIndex))
            {
                return false;
            }
            if (!HasValidSuffixBoundary(expression, closingBracketIndex))
            {
                return false;
            }
            identifierLength = closingBracketIndex - startIndex + 1;
            return true;
        }

        /// <summary>
        /// Prevents structured references such as Table1[1] from being interpreted as external workbook IDs.
        /// </summary>
        private static bool HasValidPrefixBoundary(string expression, int openingBracketIndex)
        {
            if (openingBracketIndex == 0)
            {
                return true;
            }

            char previous = expression[openingBracketIndex - 1];
            // Quoted external sheet reference: '[1]Sheet name'!A1
            if (previous == '\'')
            {
                return true;
            }
            // Table1[1], SomeName[2], etc.
            return !IsNameCharacter(previous);
        }

        /// <summary>
        /// Ensures that the numeric bracket token is followed by something that can form an external workbook reference.
        /// </summary>
        private static bool HasValidSuffixBoundary(string expression, int closingBracketIndex)
        {
            int nextIndex = closingBracketIndex + 1;
            if (nextIndex >= expression.Length)
            {
                // A bare [1] can be a structured table-column reference.
                return false;
            }
            char next = expression[nextIndex];
            // External defined name / workbook prefix: [1]!ExternalName
            if (next == '!')
            {
                return true;
            }
            // Broken external sheet reference:  [1]#REF!A1
            if (next == '#')
            {
                return true;
            }
            // The sheet or external name must immediately follow the ID.
            if (char.IsWhiteSpace(next))
            {
                return false;
            }
            switch (next)
            {
                case '"':
                case '[':
                case ']':
                case '(':
                case ')':
                case ',':
                case ';':
                case '+':
                case '-':
                case '*':
                case '/':
                case '^':
                case '&':
                case '=':
                case '<':
                case '>':
                case '%':
                case ':':
                    return false;
                default:
                    return true;
            }
        }

        /// <summary>
        /// Determines whether the passed character is a valid character for a Excel-internal name (e.g. for defined name)
        /// </summary>
        /// <param name="character">Character to check</param>
        /// <returns>True if valid, otherwise false</returns>
        private static bool IsNameCharacter(char character)
        {
            return char.IsLetterOrDigit(character) ||
                   character == '_' ||
                   character == '\\' ||
                   character == '.';
        }

        #endregion

    }
}
