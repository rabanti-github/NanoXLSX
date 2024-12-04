/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way  
 * Copyright Raphael Stoeckli © 2024
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using NanoXLSX.Shared.Exceptions;
using System.Collections.Generic;
using System.Text;
using static NanoXLSX.Shared.Enums.Styles.NumberFormatEnums;

namespace NanoXLSX.Styles
{
    /// <summary>
    /// Class representing a NumberFormat entry. The NumberFormat entry is used to define cell formats like currency or date
    /// </summary>
    public class NumberFormat : AbstractStyle
    {
        #region constants
        /// <summary>
        /// Start ID for custom number formats as constant (value 164)
        /// </summary>
        public static readonly int CUSTOMFORMAT_START_NUMBER = 164;
        /// <summary>
        /// Default format number as constant
        /// </summary>
        public static readonly FormatNumber DEFAULT_NUMBER = FormatNumber.none;

        #endregion
        private int customFormatID;
        #region enums

        #endregion

        #region privateFields
        private string customFormatCode;

        #endregion

        #region properties
        /// <summary>
        /// Gets or sets the raw custom format code in the notation of Excel. <b>The code is not escaped or un-escaped (on workbook loading)</b>
        /// </summary>
        /// <exception cref="NanoXLSX.Shared.Exceptions.FormatException">Throws a FormatException if passed value is null or empty</exception>
        /// <remarks>Currently, there is no auto-escaping applied to custom format strings. For instance, to add a white space, internally it is escaped by a backspace (\ ).
        /// To get a valid custom format code, this escaping must be applied manually, according to OOXML specs: Part 1 - Fundamentals And Markup Language Reference, Chapter 18.8.31</remarks>
        [Append]
        public string CustomFormatCode
        {
            get => customFormatCode;
            set
            {
                if (string.IsNullOrEmpty(value))
                {
                    throw new FormatException("A custom format code cannot be null or empty");
                }
                customFormatCode = value;
            }
        }
        /// <summary>
        /// Gets or sets the format number of the custom format. Must be higher or equal then predefined custom number (164) 
        /// </summary>
        /// <exception cref="NanoXLSX.Shared.Exceptions.StyleException">Throws a StyleException if the number is below the lowest possible custom number (164)</exception>
        [Append]
        public int CustomFormatID
        {
            get { return customFormatID; }
            set
            {
                if (value < CUSTOMFORMAT_START_NUMBER && !StyleRepository.Instance.ImportInProgress)
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
            customFormatCode = string.Empty;
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
            StringBuilder sb = new StringBuilder();
            sb.Append("\"NumberFormat\": {\n");
            AddPropertyAsJson(sb, "CustomFormatCode", CustomFormatCode);
            AddPropertyAsJson(sb, "CustomFormatID", CustomFormatID);
            AddPropertyAsJson(sb, "Number", Number);
            AddPropertyAsJson(sb, "HashCode", this.GetHashCode(), true);
            sb.Append("\n}");
            return sb.ToString();
        }

        /// <summary>
        /// Method to copy the current object to a new one without casting
        /// </summary>
        /// <returns>Copy of the current object without the internal ID</returns>
        public override AbstractStyle Copy()
        {
            NumberFormat copy = new NumberFormat();
            copy.customFormatCode = customFormatCode;
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
            int hashCode = 495605284;
            hashCode = hashCode * -1521134295 + EqualityComparer<string>.Default.GetHashCode(CustomFormatCode);
            hashCode = hashCode * -1521134295 + CustomFormatID.GetHashCode();
            hashCode = hashCode * -1521134295 + Number.GetHashCode();
            return hashCode;
        }

        /// <summary>
        /// Returns whether two instances are the same
        /// </summary>
        /// <param name="obj">Object to compare</param>
        /// <returns>True if this instance and the other are the same</returns>
        public override bool Equals(object obj)
        {
            return obj is NumberFormat format &&
                   CustomFormatCode == format.CustomFormatCode &&
                   CustomFormatID == format.CustomFormatID &&
                   Number == format.Number;
        }

        #endregion
    }
}
