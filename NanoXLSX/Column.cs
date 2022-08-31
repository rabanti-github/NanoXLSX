/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way  
 * Copyright Raphael Stoeckli © 2022
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using NanoXLSX.Exceptions;

namespace NanoXLSX
{
    /// <summary>
    /// Class representing a column of a worksheet
    /// </summary>
    public class Column
    {
        private int number;
        private string columnAddress;
        private float width;

        /// <summary>
        /// Column address (A to XFD)
        /// </summary>
        public string ColumnAddress
        {
            get { return columnAddress; }
            set
            {
                if (string.IsNullOrEmpty(value))
                {
                    throw new RangeException("The passed address was null or empty");
                }
                number = Cell.ResolveColumn(value);
                columnAddress = value.ToUpper();
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
        public float Width {
            get { return width; }
            set
            {
                if (value < Worksheet.MIN_COLUMN_WIDTH || value > Worksheet.MAX_COLUMN_WIDTH)
                {
                    throw new RangeException("The passed column width is out of range (" + Worksheet.MIN_COLUMN_WIDTH + " to " + Worksheet.MAX_COLUMN_WIDTH + ")");
                }
                width = value;
            }
        }
        

        /// <summary>
        /// Default constructor (private, since not valid without address)
        /// </summary>
        private Column()
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

        /// <summary>
        /// Creates a deep copy of this column
        /// </summary>
        /// <returns>Copy of this column</returns>
        internal Column Copy()
        {
            Column copy = new Column();
            copy.IsHidden = this.IsHidden;
            copy.Width = this.width;
            copy.HasAutoFilter = this.HasAutoFilter;
            copy.columnAddress = this.columnAddress;
            copy.number = this.number;
            return copy;
        }

    }
}
