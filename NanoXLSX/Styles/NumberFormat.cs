/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way  
 * Copyright Raphael Stoeckli © 2020
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System.Text;

namespace Styles
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
        public const int CUSTOMFORMAT_START_NUMBER = 124;
        #endregion

        #region enums
        /// <summary>
        /// Enum for predefined number formats
        /// </summary>
        /// <remarks>There are other predefined formats (e.g. 43 and 44) that are not listed. The declaration of such formats is done in the number formats section of the style document, whereas the officially listed ones are implicitly used and not declared in the style document</remarks>
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
        #endregion

        #region properties
        /// <summary>
        /// Gets or sets the custom format code in the notation of Excel
        /// </summary>
        public string CustomFormatCode { get; set; }
        /// <summary>
        /// Gets or sets the format number of the custom format. Must be higher or equal then predefined custom number (164) 
        /// </summary>
        public int CustomFormatID { get; set; }
        /// <summary>
        /// Gets whether the number format is a custom format (higher or equals 164). If true, the format is custom
        /// </summary>
        [Append(Ignore = true)]
        public bool IsCustomFormat
        {
            get
            {
                if (Number == FormatNumber.custom) { return true; }
                else { return false; }
            }
        }
        /// <summary>
        /// Gets or sets the format number. Set this to custom (164) in case of custom number formats
        /// </summary>
        public FormatNumber Number { get; set; }
        #endregion

        #region constructors
        /// <summary>
        /// Default constructor
        /// </summary>
        public NumberFormat()
        {
            Number = FormatNumber.none;
            CustomFormatCode = string.Empty;
            CustomFormatID = CUSTOMFORMAT_START_NUMBER;
        }
        #endregion

        #region methods

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
            int p = 251;
            int r = 1;
            r *= p + this.CustomFormatCode.GetHashCode();
            r *= p + this.CustomFormatID;
            r *= p + (int)this.Number;
            return r;
        }

        #endregion
    }
}