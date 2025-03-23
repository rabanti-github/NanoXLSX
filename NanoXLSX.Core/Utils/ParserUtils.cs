/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2025
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System;
using System.Globalization;
using System.Linq;

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
        /// Constant for number conversion. The invariant culture (represents mostly the US numbering scheme) ensures that no culture-specific 
        /// punctuations are used when converting numbers to strings, This is especially important for OOXML number values.
        /// See also: <a href="https://docs.microsoft.com/en-us/dotnet/api/system.globalization.cultureinfo.invariantculture?view=net-5.0">
        /// https://docs.microsoft.com/en-us/dotnet/api/system.globalization.cultureinfo.invariantculture?view=net-5.0</a>
        /// </summary>
        public static readonly CultureInfo INVARIANT_CULTURE = CultureInfo.InvariantCulture;

        #endregion

        /// <summary>
        /// Transforms a string to upper case with null check and invariant culture
        /// </summary>
        /// <param name="input">String to transform</param>
        /// <returns>Upper case string</returns>
        public static string ToUpper(string input)
        {
            return !string.IsNullOrEmpty(input) ? input.ToUpper(INVARIANT_CULTURE) : input;
        }

        /// <summary>
        /// Transforms an integer to an invariant sting
        /// </summary>
        /// <param name="input">Integer to transform</param>
        /// <returns>Integer as string</returns>
        public static string ToString(int input)
        {
            return input.ToString("G", INVARIANT_CULTURE);
        }

        /// <summary>
        /// Transforms a float to an invariant sting
        /// </summary>
        /// <param name="input">Float to transform</param>
        /// <returns>Float as string</returns>
        public static string ToString(float input)
        {
            return input.ToString("G", INVARIANT_CULTURE);
        }

        /// <summary>
        /// Transforms a byte to an invariant sting
        /// </summary>
        /// <param name="input">Byte to transform</param>
        /// <returns>Byte as string</returns>
        public static string ToString(byte input)
        {
            return input.ToString("G", INVARIANT_CULTURE);
        }

        /// <summary>
        /// Transforms a sbyte to an invariant sting
        /// </summary>
        /// <param name="input">Sbyte to transform</param>
        /// <returns>Byte as string</returns>
        public static string ToString(sbyte input)
        {
            return input.ToString("G", INVARIANT_CULTURE);
        }

        /// <summary>
        /// Transforms a double to an invariant sting
        /// </summary>
        /// <param name="input">Double to transform</param>
        /// <returns>Double as string</returns>
        public static string ToString(double input)
        {
            return input.ToString("G", INVARIANT_CULTURE);
        }

        /// <summary>
        /// Transforms a decimal to an invariant sting
        /// </summary>
        /// <param name="input">Decimal to transform</param>
        /// <returns>Decimal as string</returns>
        public static string ToString(decimal input)
        {
            return input.ToString("G", INVARIANT_CULTURE);
        }

        /// <summary>
        /// Transforms a uint to an invariant sting
        /// </summary>
        /// <param name="input">Uint to transform</param>
        /// <returns>Uint as string</returns>
        public static string ToString(uint input)
        {
            return input.ToString("G", INVARIANT_CULTURE);
        }

        /// <summary>
        /// Transforms a long to an invariant sting
        /// </summary>
        /// <param name="input">Long to transform</param>
        /// <returns>Long as string</returns>
        public static string ToString(long input)
        {
            return input.ToString("G", INVARIANT_CULTURE);
        }

        /// <summary>
        /// Transforms a ulong to an invariant sting
        /// </summary>
        /// <param name="input">Ulong to transform</param>
        /// <returns>Ulong as string</returns>
        public static string ToString(ulong input)
        {
            return input.ToString("G", INVARIANT_CULTURE);
        }

        /// <summary>
        /// Transforms a short to an invariant sting
        /// </summary>
        /// <param name="input">Short to transform</param>
        /// <returns>Short as string</returns>
        public static string ToString(short input)
        {
            return input.ToString("G", INVARIANT_CULTURE);
        }

        /// <summary>
        /// Transforms a ushort to an invariant sting
        /// </summary>
        /// <param name="input">Ushort to transform</param>
        /// <returns>Ushort as string</returns>
        public static string ToString(ushort input)
        {
            return input.ToString("G", INVARIANT_CULTURE);
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
        public static float ParseFloat(string rawValue)
        {
            return float.Parse(rawValue, INVARIANT_CULTURE);
        }

        /// <summary>
        /// Parses an int independent from the culture info of the host
        /// </summary>
        /// <param name="rawValue">Raw number as string</param>
        /// <returns>Parsed int</returns>
        public static int ParseInt(string rawValue)
        {
            return int.Parse(rawValue, NumberStyles.Any, INVARIANT_CULTURE);
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
        /// <param name="rawValue">Raw number as string</param>
        /// <param name="parsedValue">Parsed int</param>
        /// <returns>True, if the parsing was successful</returns>
        public static bool TryParseInt(string rawValue, out int parsedValue)
        {
            return int.TryParse(rawValue, NumberStyles.Integer, INVARIANT_CULTURE, out parsedValue);
        }

        /// <summary>
        /// Tries to parse an unsigned int (uint) independent from the culture info of the host
        /// </summary>
        /// <param name="rawValue">Raw number as string</param>
        /// <param name="parsedValue">Parsed uint</param>
        /// <returns>True, if the parsing was successful</returns>
        public static bool TryParseUint(string rawValue, out uint parsedValue)
        {
            return uint.TryParse(rawValue, NumberStyles.Integer, INVARIANT_CULTURE, out parsedValue);
        }

        /// <summary>
        /// Tries to parse a long independent from the culture info of the host
        /// </summary>
        /// <param name="rawValue">Raw number as string</param>
        /// <param name="parsedValue">Parsed long</param>
        /// <returns>True, if the parsing was successful</returns>
        public static bool TryParseLong(string rawValue, out long parsedValue)
        {
            return long.TryParse(rawValue, NumberStyles.Integer, INVARIANT_CULTURE, out parsedValue);
        }

        /// <summary>
        /// Tries to parse an unsigned long (ulong) independent from the culture info of the host
        /// </summary>
        /// <param name="rawValue">Raw number as string</param>
        /// <param name="parsedValue">Parsed ulong</param>
        /// <returns>True, if the parsing was successful</returns>
        public static bool TryParseUlong(string rawValue, out ulong parsedValue)
        {
            return ulong.TryParse(rawValue, NumberStyles.Integer, INVARIANT_CULTURE, out parsedValue);
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
        /// <param name="rawvalue">Raw number as string</param>
        /// <param name="parsedValue">Parsed decimal</param>
        /// <returns>True, if the parsing was successful</returns>
        public static bool TryParseDecimal(string rawvalue, out decimal parsedValue)
        {
            return decimal.TryParse(rawvalue, NumberStyles.Float, INVARIANT_CULTURE, out parsedValue);
        }

        /// <summary>
        /// Tries to parse a double (with any parsing style) independent from the culture info of the host
        /// </summary>
        /// <param name="rawValue">Raw number as string</param>
        /// <param name="parsedValue">Parsed double</param>
        /// <returns>True, if the parsing was successful</returns>
        public static bool TryParseDouble(string rawValue, out double parsedValue)
        {
            return double.TryParse(rawValue, NumberStyles.Any, INVARIANT_CULTURE, out parsedValue);
        }
    }
}
