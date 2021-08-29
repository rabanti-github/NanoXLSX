/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way  
 * Copyright Raphael Stoeckli © 2021
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using NanoXLSX.Exceptions;

namespace NanoXLSX.Styles
{
    /// <summary>
    /// Class representing a NumberFormat entry. The NumberFormat entry is used to define cell formats like currency or date
    /// </summary>
    public class NumberFormat : AbstractStyle
    {
        #region constants
        /// <summary>
        /// Start ID for custom number formats as constant
        /// </summary>
        public static readonly int CUSTOMFORMAT_START_NUMBER = 164;
        /// <summary>
        /// Default format number as constant
        /// </summary>
        public static readonly FormatNumber DEFAULT_NUMBER = FormatNumber.none;

        #endregion
        private int customFormatID;
        #region enums
        /// <summary>
        /// Enum for predefined number formats
        /// </summary>
        /// <remarks>There are other predefined formats (e.g. 43 and 44) that are not listed. The declaration of such formats is done in the number formats section of the style document, 
        /// whereas the officially listed ones are implicitly used and not declared in the style document</remarks>
        public enum FormatNumber
        {
            /// <summary>No format / Default</summary>
            none = 0,
            /// <summary>Format: 0</summary>
            format_1 = 1,
            /// <summary>Format: 0.00</summary>
            format_2 = 2,
            /// <summary>Format: #,##0</summary>
            format_3 = 3,
            /// <summary>Format: #,##0.00</summary>
            format_4 = 4,
            /// <summary>Format: $#,##0_);($#,##0)</summary>
            format_5 = 5,
            /// <summary>Format: $#,##0_);[Red]($#,##0)</summary>
            format_6 = 6,
            /// <summary>Format: $#,##0.00_);($#,##0.00)</summary>
            format_7 = 7,
            /// <summary>Format: $#,##0.00_);[Red]($#,##0.00)</summary>
            format_8 = 8,
            /// <summary>Format: 0%</summary>
            format_9 = 9,
            /// <summary>Format: 0.00%</summary>
            format_10 = 10,
            /// <summary>Format: 0.00E+00</summary>
            format_11 = 11,
            /// <summary>Format: # ?/?</summary>
            format_12 = 12,
            /// <summary>Format: # ??/??</summary>
            format_13 = 13,
            /// <summary>Format: m/d/yyyy</summary>
            format_14 = 14,
            /// <summary>Format: d-mmm-yy</summary>
            format_15 = 15,
            /// <summary>Format: d-mmm</summary>
            format_16 = 16,
            /// <summary>Format: mmm-yy</summary>
            format_17 = 17,
            /// <summary>Format: mm AM/PM</summary>
            format_18 = 18,
            /// <summary>Format: h:mm:ss AM/PM</summary>
            format_19 = 19,
            /// <summary>Format: h:mm</summary>
            format_20 = 20,
            /// <summary>Format: h:mm:ss</summary>
            format_21 = 21,
            /// <summary>Format: m/d/yyyy h:mm</summary>
            format_22 = 22,
            /// <summary>Format: #,##0_);(#,##0)</summary>
            format_37 = 37,
            /// <summary>Format: #,##0_);[Red](#,##0)</summary>
            format_38 = 38,
            /// <summary>Format: #,##0.00_);(#,##0.00)</summary>
            format_39 = 39,
            /// <summary>Format: #,##0.00_);[Red](#,##0.00)</summary>
            format_40 = 40,
            /// <summary>Format: mm:ss</summary>
            format_45 = 45,
            /// <summary>Format: [h]:mm:ss</summary>
            format_46 = 46,
            /// <summary>Format: mm:ss.0</summary>
            format_47 = 47,
            /// <summary>Format: ##0.0E+0</summary>
            format_48 = 48,
            /// <summary>Format: #</summary>
            format_49 = 49,
            /// <summary>Custom Format (ID 164 and higher)</summary>
            custom = 164,
        }

        /// <summary>
        /// Range or validity of the format number
        /// </summary>
        public enum FormatRange
        {
            /// <summary>
            /// Format from 0 to 164 (with gaps)
            /// </summary>
            defined_format,
            /// <summary>
            /// Custom defined formats from 165 and higher. Although 164 is already custom, it is still defined as enum value
            /// </summary>
            custom_format,
            /// <summary>
            /// Probably invalid format numbers (e.g. negative value)
            /// </summary>
            invalid,
            /// <summary>
            /// Values between 0 and 164 that are not defined as enum value. This may be caused by changes of the OOXML specifications or Excel versions that have encoded loaded files
            /// </summary>
            undefined,
        }
        #endregion

        #region privateFields
        #endregion

        #region properties
        /// <summary>
        /// Gets or sets the custom format code in the notation of Excel
        /// </summary>
        [Append]
        public string CustomFormatCode { get; set; }
        /// <summary>
        /// Gets or sets the format number of the custom format. Must be higher or equal then predefined custom number (164) 
        /// </summary>
        /// <exception cref="Exceptions.StyleException">Throws a StyleException if the number is below the lowest possible custom number (164)</exception>
        [Append]
        public int CustomFormatID
        {
            get { return customFormatID; }
            set
            {
                if (value < CUSTOMFORMAT_START_NUMBER)
                {
                    throw new StyleException("The number '" + value + "' is not a valid custom format ID. Must be at least " + CUSTOMFORMAT_START_NUMBER);
                }
                customFormatID = value;
            }
        }
        /// <summary>
        /// Gets whether the number format is a custom format (higher or equals 164). If true, the format is custom
        /// </summary>
        [Append(Ignore = true)]
        public bool IsCustomFormat
        {
            get
            {
                if (Number == FormatNumber.custom)
                { return true; }
                else { return false; }
            }
        }
        /// <summary>
        /// Gets or sets the format number. Set this to custom (164) in case of custom number formats
        /// </summary>
        [Append]
        public FormatNumber Number { get; set; }
        #endregion

        #region constructors
        /// <summary>
        /// Default constructor
        /// </summary>
        public NumberFormat()
        {
            Number = DEFAULT_NUMBER;
            CustomFormatCode = string.Empty;
            CustomFormatID = CUSTOMFORMAT_START_NUMBER;
        }
        #endregion

        #region methods

        /// <summary>
        /// Determines whether a defined style format number represents a date (or date and time)
        /// </summary>
        /// <param name="number">Format number to check</param>
        /// <returns>True if the format represents a date, otherwise false</returns>
        /// <remarks>Custom number formats (higher than 164), as well as not officially defined numbers (below 164) are currently not considered during the check and will return false</remarks>
        public static bool IsDateFormat(FormatNumber number)
        {
            switch (number)
            {
                case FormatNumber.format_14:
                case FormatNumber.format_15:
                case FormatNumber.format_16:
                case FormatNumber.format_17:
                case FormatNumber.format_22:
                    return true;
                default:
                    return false;
            }
        }

        /// <summary>
        /// Determines whether a defined style format number represents a time)
        /// </summary>
        /// <param name="number">Format number to check</param>
        /// <returns>True if the format represents a time, otherwise false</returns>
        /// <remarks>Custom number formats (higher than 164), as well as not officially defined numbers (below 164) are currently not considered during the check and will return false</remarks>
        public static bool IsTimeFormat(FormatNumber number)
        {
            switch (number)
            {
                case FormatNumber.format_18:
                case FormatNumber.format_19:
                case FormatNumber.format_20:
                case FormatNumber.format_21:
                case FormatNumber.format_45:
                case FormatNumber.format_46:
                case FormatNumber.format_47:
                    return true;
                default:
                    return false;
            }
        }

        /// <summary>
        /// Tries to parse registered format numbers. If the parsing fails, it is assumed that the number is a custom format number (164 or higher) and 'custom' is returned 
        /// </summary>
        /// <param name="number">Raw number to parse</param>
        /// <param name="formatNumber">Out parameter with the parsed format enum value. If parsing failed, 'custom' will be returned</param>
        /// <returns>Format range. Will return 'invalid' if out of any range (e.g. negative value)</returns>
        public static FormatRange TryParseFormatNumber(int number, out FormatNumber formatNumber)
        {

            bool isDefined = System.Enum.IsDefined(typeof(FormatNumber), number);
            if (isDefined)
            {
                formatNumber = (FormatNumber)number;
                return FormatRange.defined_format;
            }
            if (number < 0)
            {
                formatNumber = FormatNumber.none;
                return FormatRange.invalid;
            }
            else if (number > 0 && number < CUSTOMFORMAT_START_NUMBER)
            {
                formatNumber = FormatNumber.none;
                return FormatRange.undefined;
            }
            else
            {
                formatNumber = FormatNumber.custom;
                return FormatRange.custom_format;
            }
        }

        /// <summary>
        /// Override toString method
        /// </summary>
        /// <returns>String of a class</returns>
        public override string ToString()
        {
            return "NumberFormat:" + this.GetHashCode();
        }

        /// <summary>
        /// Method to copy the current object to a new one without casting
        /// </summary>
        /// <returns>Copy of the current object without the internal ID</returns>
        public override AbstractStyle Copy()
        {
            NumberFormat copy = new NumberFormat();
            copy.CustomFormatCode = CustomFormatCode;
            copy.CustomFormatID = CustomFormatID;
            copy.Number = Number;
            return copy;
        }

        /// <summary>
        /// Method to copy the current object to a new one with casting
        /// </summary>
        /// <returns>Copy of the current object without the internal ID</returns>
        public NumberFormat CopyNumberFormat()
        {
            return (NumberFormat)Copy();
        }

        /// <summary>
        /// Returns a hash code for this instance.
        /// </summary>
        /// <returns>
        /// A hash code for this instance, suitable for use in hashing algorithms and data structures like a hash table. 
        /// </returns>
        public override int GetHashCode()
        {
            const int p = 251;
            int r = 1;
            r *= p + this.CustomFormatCode.GetHashCode();
            r *= p + this.CustomFormatID;
            r *= p + (int)this.Number;
            return r;
        }

        #endregion
    }
}
