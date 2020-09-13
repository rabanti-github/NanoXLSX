/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way  
 * Copyright Raphael Stoeckli © 2020
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

namespace NanoXLSX
{
    /// <summary>
    /// Class representing a column of a worksheet
    /// </summary>
    public class Column
    {
        private int number;
        private string columnAddress;

        /// <summary>
        /// Column address (A to XFD)
        /// </summary>
        public string ColumnAddress
        {
            get { return columnAddress; }
            set
            {
                number = Cell.ResolveColumn(value);
                columnAddress = value;
            }
        }

        /// <summary>
        /// If true, the column has auto filter applied, otherwise not
        /// </summary>
        public bool HasAutoFilter { get; set; }
        /// <summary>
        /// If true, the column is hidden, otherwise visible
        /// </summary>
        public bool IsHidden { get; set; }

        /// <summary>
        /// Column number (0 to 16383)
        /// </summary>
        public int Number
        {
            get { return number; }
            set
            {
                columnAddress = Cell.ResolveColumnAddress(value);
                number = value;
            }
        }

        /// <summary>
        /// Width of the column
        /// </summary>
        public float Width { get; set; }

        /// <summary>
        /// Default constructor
        /// </summary>
        public Column()
        {
            Width = Worksheet.DEFAULT_COLUMN_WIDTH;
        }

        /// <summary>
        /// Constructor with column number
        /// </summary>
        /// <param name="columnCoordinate">Column number (zero-based, 0 to 16383)</param>
        public Column(int columnCoordinate) : this()
        {
            Number = columnCoordinate;
        }

        /// <summary>
        /// Constructor with column address
        /// </summary>
        /// <param name="columnAddress">Column address (A to XFD)</param>
        public Column(string columnAddress) : this()
        {
            ColumnAddress = columnAddress;
        }

    }
}
