/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way  
 * Copyright Raphael Stoeckli © 2025
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System.Collections.Generic;
using System.Globalization;
using NanoXLSX.Exceptions;
using NanoXLSX.Interfaces.Plugin;

namespace NanoXLSX
{
    /// <summary>
    /// The reader options define global rules, applied when loading a worksheet. The options are mainly to override particular cell types (e.g. interpretation of dates as numbers)
    /// </summary>
    public class ReaderOptions : IOptions
    {

        /// <summary>
        /// Default format if DateTime values are cast to strings
        /// </summary>
        public const string DEFAULT_DATE_TIME_FORMAT = "yyyy-MM-dd HH:mm:ss";

        /// <summary>
        /// Default format if TimeSpan values are cast to strings
        /// </summary>
        public const string DEFAULT_TIMESPAN_FORMAT = "hh\\:mm\\:ss";

        /// <summary>
        /// Default culture info instance (invariant culture) used for date and time parsing, if no custom culture info is defined
        /// </summary>
        public static readonly CultureInfo DEFAULT_CULTURE_INFO = CultureInfo.InvariantCulture;

        /// <summary>
        /// Global conversion types to enforce during the load process. All types other than <see cref="GlobalType.Default" /> will override defined <see cref="ColumnType">Column types</a>
        /// </summary>
        public enum GlobalType
        {
            /// <summary>
            /// No global strategy. All numbers are tried to be cast to the most suitable types
            /// </summary>
            Default,
            /// <summary>
            /// All numbers are cast to doubles
            /// </summary>
            AllNumbersToDouble,
            /// <summary>
            /// All numbers are cast to decimal
            /// </summary>
            AllNumbersToDecimal,
            /// <summary>
            /// All numbers are cast to integers. Floating point numbers will be rounded (commercial rounding) to the nearest integers
            /// </summary>
            AllNumbersToInt,
            /// <summary>
            /// Every cell is cast to a string
            /// </summary>
            EverythingToString
        }

        /// <summary>
        /// Column types to enforce during the read process. The types are tried to be applied on all cells of a particular column
        /// </summary>
        public enum ColumnType
        {
            /// <summary>
            /// Cells are tried to be imported as numbers (automatic determination of numeric type)
            /// </summary>
            Numeric,
            /// <summary>
            /// Cells are tried to be imported as numbers (enforcing double)
            /// </summary>
            Double,
            /// <summary>
            /// Cells are tried to be imported as numbers (enforcing decimal)
            /// </summary>
            Decimal,
            /// <summary>
            /// Cells are tried to be imported as dates (DateTime). See also  <see cref="DateTimeFormat"/>, <see cref="TimeSpanFormat"/> and <see cref="TemporalCultureInfo"/>
            /// </summary>
            Date,
            /// <summary>
            /// Cells are tried to be imported as times (TimeSpan)
            /// </summary>
            Time,
            /// <summary>
            /// Cells are tried to be imported as bools
            /// </summary>
            Bool,
            /// <summary>
            /// Cells are all imported as strings, using the ToString() method
            /// </summary>
            String
        }

        /// <summary>
        /// If true, date or time values (default format number 14 or 21) will be interpreted as numeric values globally. 
        /// This option overrules possible column options, defined by <see cref="AddEnforcedColumn(int, ColumnType)"/>.
        /// </summary>
        public bool EnforceDateTimesAsNumbers { get; set; }

        /// <summary>
        /// If true, phonetic characters (like ruby characters / Furigana / Zhuyin fuhao) in strings are added in brackets after the transcribed symbols. By default, phonetic characters are removed from strings.
        /// </summary>
        /// \remark <remarks>This option is not applicable to specific rows or a start column (applied globally)</remarks>
        public bool EnforcePhoneticCharacterImport { get; set; }

        /// <summary>
        /// If true, empty cells will be interpreted as type of string with an empty value. If false, the type will be Empty and the value null
        /// </summary>
        public bool EnforceEmptyValuesAsString { get; set; }

        /// <summary>
        /// If true, invalid data, like column widths or row height that are out of range, will cause an exception when such a workbook is loaded. Tho option is inactive by default (tolerant reader mode)
        /// </summary>
        public bool EnforceStrictValidation { get; set; }

        /// <summary>
        /// Global strategy to handle cell values. The default will not enforce any general casting, beside defined values of <see cref="EnforceDateTimesAsNumbers" />, <see cref="EnforceEmptyValuesAsString" /> and <see cref="EnforcedColumnTypes" /> 
        /// </summary>
        public GlobalType GlobalEnforcingType { get; set; } = GlobalType.Default;


        /// <summary>
        /// Type enforcing rules during the read process for particular columns
        /// </summary>
        public Dictionary<int, ColumnType> EnforcedColumnTypes { get; private set; } = new Dictionary<int, ColumnType>();

        /// <summary>
        /// The row number (zero-based) where enforcing rules are started to be applied. This is, for instance, to prevent enforcing types in a header row. Any enforcing rule is skipped until this row number is reached
        /// </summary>
        public int EnforcingStartRowNumber { get; set; } = 0;

        /// <summary>
        /// Format if DateTime values are cast to strings or DateTime objects are parsed from strings. If null or empty, parsing will be tried with 'best effort', according to <see cref="System.DateTime.Parse(string)"> System.DateTime.Parse(string)</see>. 
        /// See also  <see cref="TemporalCultureInfo"/>
        /// </summary>
        public string DateTimeFormat { get; set; } = DEFAULT_DATE_TIME_FORMAT;

        /// <summary>
        /// Format if TimeSpan values are cast to strings
        /// </summary>
        /// \remark <remarks>The separators like period or semicolon must be escaped by backslashes. See: <a href="https://docs.microsoft.com/en-us/dotnet/standard/base-types/custom-timespan-format-strings"/></remarks>
        public string TimeSpanFormat { get; set; } = DEFAULT_TIMESPAN_FORMAT;

        /// <summary>
        /// Culture info instance, used to parse DateTime or TimeSpan objects from strings. If null, parsing will be tried with 'best effort', according to <see cref="System.DateTime.Parse(string)">System.DateTime.Parse(string)</see> <see cref="System.DateTime.Parse(string)">System.DateTime.Parse(string)</see>.
        /// See also  <see cref="DateTimeFormat"/> and <see cref="TimeSpanFormat"/>
        /// </summary>
        public CultureInfo TemporalCultureInfo { get; set; } = DEFAULT_CULTURE_INFO;

        /// <summary>
        /// If set to true, worksheet or workbook protection passwords with unknown / not supported algorithms will be ignored (password hash may not be read). 
        /// Otherwise, a <see cref="NotSupportedContentException"/> will be thrown. Default is false 
        /// </summary>
        public bool IgnoreNotSupportedPasswordAlgorithms { get; set; }

        /// <summary>
        /// Adds a type enforcing rule to the passed column address
        /// </summary>
        /// <param name="columnAddress">Column address (A to XFD)</param>
        /// <param name="type">Type to be enforced on the column</param>
        public void AddEnforcedColumn(string columnAddress, ColumnType type)
        {
            this.EnforcedColumnTypes.Add(Cell.ResolveColumn(columnAddress), type);
        }

        /// <summary>
        /// Adds a type enforcing rule to the passed column number (zero-based)
        /// </summary>
        /// <param name="columnNumber">Column number (0-16383)</param>
        /// <param name="type">Type to be enforced on the column</param>
        public void AddEnforcedColumn(int columnNumber, ColumnType type)
        {
            this.EnforcedColumnTypes.Add(columnNumber, type);
        }
    }
}
