/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way  
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using NanoXLSX.Exceptions;
using NanoXLSX.Styles;
using NanoXLSX.Utils;

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
        private Style defaultColumnStyle;

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
                columnAddress = ParserUtils.ToUpper(value);
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
        public float Width
        {
            get { return width; }
            set
            {
                if (value < Worksheet.MinColumnWidth || value > Worksheet.MaxColumnWidth)
                {
                    throw new RangeException("The passed column width is out of range (" + Worksheet.MinColumnWidth + " to " + Worksheet.MaxColumnWidth + ")");
                }
                width = value;
            }
        }

        /// <summary>
        /// Gets the default style of the column
        /// </summary>
        public Style DefaultColumnStyle
        {
            get { return this.defaultColumnStyle; }
        }

        /// <summary>
        /// Sets the default style of the column
        /// </summary>
        /// <param name="defaultColumnStyle">Style to assign as default column style. Can be null (to clear)</param>
        /// <param name="unmanaged">Internally used: If true, the style repository is not invoked and only the style object of the cell is updated. Do not use!</param>
        /// <returns>If the passed style already exists in the repository, the existing one will be returned, otherwise the passed one</returns>
        public Style SetDefaultColumnStyle(Style defaultColumnStyle, bool unmanaged = false)
        {
            if (defaultColumnStyle == null)
            {
                this.defaultColumnStyle = null;
                return null;
            }
            if (unmanaged)
            {
                this.defaultColumnStyle = defaultColumnStyle;
            }
            else
            {
                this.defaultColumnStyle = StyleRepository.Instance.AddStyle(defaultColumnStyle);
            }
            return this.defaultColumnStyle;
        }

        /// <summary>
        /// Default constructor (private, since not valid without address)
        /// </summary>
        private Column()
        {
            Width = Worksheet.DefaultWorksheetColumnWidth;
            defaultColumnStyle = null;
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
            Column copy = new Column
            {
                IsHidden = this.IsHidden,
                Width = this.width,
                HasAutoFilter = this.HasAutoFilter,
                columnAddress = this.columnAddress,
                number = this.number,
                defaultColumnStyle = this.defaultColumnStyle
            };
            return copy;
        }

    }
}
