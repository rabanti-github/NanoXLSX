/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way  
 * Copyright Raphael Stoeckli © 2022
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System.Collections.Generic;

namespace NanoXLSX
{
    /// <summary>
    /// The import options define global rules to import worksheets. The options are mainly to override particular cell types (e.g. interpretation of dates as numbers)
    /// </summary>
    public class ImportOptions
    {
        /// <summary>
        /// Column types to enforce during the import
        /// </summary>
        public enum ColumnType
        {
            /// <summary>
            /// Cells are tried to be imported as numbers (double)
            /// </summary>
            Numeric,
            /// <summary>
            /// Cells are tried to be imported as dates (DateTime)
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
        /// Type enforcing rules during import for particular columns
        /// </summary>
        public Dictionary<int, ColumnType> EnforcedColumnTypes { get; private set; } = new Dictionary<int, ColumnType>();

        /// <summary>
        /// The row number (zero-based) where enforcing rules are started to be applied. This is, for instance, to prevent enforcing in a header row
        /// </summary>
        public int EnforcingStartRowNumber { get; set; }

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
