/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2022
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using NanoXLSX.Shared.Interfaces;

namespace NanoXLSX.Shared.Enums.Styles
{
    /// <summary>
    /// Class providing shared enums used by implementations of the <see cref="INumberFormat"/> interface // TODO: Define interface
    /// </summary>
    public static class NumberFormatEnums
    {
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
    }
}
