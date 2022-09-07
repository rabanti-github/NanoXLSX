/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way  
 * Copyright Raphael Stoeckli © 2022
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System;

namespace NanoXLSX
{
    /// <summary>
    /// Struct representing the cell address as column and row (zero based)
    /// </summary>
    public struct Address : IEquatable<Address>, IComparable<Address>
    {
        /// <summary>
        /// Column number (zero based)
        /// </summary>
        public int Column;
        /// <summary>
        /// Row number (zero based)
        /// </summary>
        public int Row;

        /// <summary>
        /// Referencing type of the address
        /// </summary>
        public Cell.AddressType Type;

        /// <summary>
        /// Constructor with row and column as arguments.  The referencing type of the address is default (e.g. 'C20')
        /// </summary>
        /// <param name="column">Column number (zero based)</param>
        /// <param name="row">Row number (zero based)</param>
        public Address(int column, int row) : this(column, row, Cell.AddressType.Default)
        {
            // No actions
        }

        /// <summary>
        /// Constructor with row and column as arguments. All referencing modifiers ($) are ignored and only the defined referencing type considered
        /// </summary>
        /// <param name="column">Column number (zero based)</param>
        /// <param name="row">Row number (zero based)</param>
        /// <param name="type">Referencing type of the address</param>
        public Address(int column, int row, Cell.AddressType type)
        {
            Cell.ValidateColumnNumber(column);
            Cell.ValidateRowNumber(row);
            Column = column;
            Row = row;
            Type = type;
        }

        /// <summary>
        /// Constructor with address as string. If no referencing modifiers ($) are defined, the address is of referencing type default (e.g. 'C23')
        /// </summary>
        /// <param name="address">Address string (e.g. '$B$12')</param>
        public Address(string address)
        {
            Cell.ResolveCellCoordinate(address, out Column, out Row, out Type);
        }

        /// <summary>
        /// Constructor with address as string. All referencing modifiers ($) are ignored and only the defined referencing type considered
        /// </summary>
        /// <param name="address">Address string (e.g. 'B12')</param>
        /// <param name="type">Referencing type of the address</param>
        public Address(string address, Cell.AddressType type)
        {
            Type = type;
            Cell.ResolveCellCoordinate(address, out Column, out Row);
        }

        /// <summary>
        /// Returns the combined Address
        /// </summary>
        /// <returns>Address as string in the format A1 - XFD1048576</returns>
        public string GetAddress()
        {
            return Cell.ResolveCellAddress(Column, Row, Type);
        }

        /// <summary>
        /// Gets the column address (A - XFD)
        /// </summary>
        /// <returns>Column address as letter(s)</returns>
        public string GetColumn()
        {
            return Cell.ResolveColumnAddress(Column);
        }

        /// <summary>
        /// Overwritten ToString method
        /// </summary>
        /// <returns>Returns the cell address (e.g. 'A15')</returns>
        public override string ToString()
        {
            return GetAddress();
        }

        /// <summary>
        /// Compares two addresses whether they are equal
        /// </summary>
        /// <param name="o"> Other address</param>
        /// <returns>True if equal</returns>
        public bool Equals(Address o)
        {
            if (Row == o.Row && Column == o.Column && Type == o.Type)
            { return true; }

            return false;
        }

        /// <summary>
        /// Compares two addresses using the column and row numbers
        /// </summary>
        /// <param name="other"> Other address</param>
        /// <returns>-1 if the other address is greater, 0 if equal and 1 if smaller</returns>
        public int CompareTo(Address other)
        {
            long thisCoordinate = (long)Column * (long)Worksheet.MAX_ROW_NUMBER + Row;
            long otherCoordinate = (long)other.Column * (long)Worksheet.MAX_ROW_NUMBER + other.Row;
            return thisCoordinate.CompareTo(otherCoordinate);
        }

        /// <summary>
        /// Creates a (dereferenced, if applicable) deep copy of this address
        /// </summary>
        /// <returns>Copy of this range</returns>
        internal Address Copy()
        {
            return new Address(this.Column, this.Row, this.Type);
        }
    }

}
