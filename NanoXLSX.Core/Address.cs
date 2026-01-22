/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way  
 * Copyright Raphael Stoeckli © 2026
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
        private readonly int column;
        private readonly int row;
        private readonly Cell.AddressType type;

        /// <summary>
        /// Column number (zero based)
        /// </summary>
        public int Column { get => column; }
        /// <summary>
        /// Row number (zero based)
        /// </summary>
        public int Row { get => row; }

        /// <summary>
        /// Referencing type of the address
        /// </summary>
        public Cell.AddressType Type { get => type; }

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
            this.column = column;
            this.row = row;
            this.type = type;
        }

        /// <summary>
        /// Constructor with address as string. If no referencing modifiers ($) are defined, the address is of referencing type default (e.g. 'C23')
        /// </summary>
        /// <param name="address">Address string (e.g. '$B$12')</param>
        public Address(string address)
        {
            Cell.ResolveCellCoordinate(address, out this.column, out this.row, out this.type);
        }

        /// <summary>
        /// Constructor with address as string. All referencing modifiers ($) are ignored and only the defined referencing type considered
        /// </summary>
        /// <param name="address">Address string (e.g. 'B12')</param>
        /// <param name="type">Referencing type of the address</param>
        public Address(string address, Cell.AddressType type)
        {
            this.type = type;
            Cell.ResolveCellCoordinate(address, out this.column, out this.row);
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
        /// <param name="other"> Other address</param>
        /// <returns>True if equal</returns>
        public bool Equals(Address other)
        {
            if (Row == other.Row && Column == other.Column && Type == other.Type)
            { return true; }

            return false;
        }

        /// <summary>
        /// Compares two objects whether they are addresses and equal
        /// </summary>
        /// <param name="obj"> Other address</param>
        /// <returns>True if not null, of the same type and equal</returns>
        public override bool Equals(object obj)
        {
            if (!(obj is Address))
            {
                return false;
            }
            return Equals((Address)obj);
        }

        /// <summary>
        /// Gets the hash code based on the string representation of the address
        /// </summary>
        /// <returns>Hash code of the address</returns>
        public override int GetHashCode()
        {
            return ToString().GetHashCode();
        }

        // Operator overloads
        /// <summary>
        /// Determines whether two <see cref="Address"/> instances are equal.
        /// </summary>
        /// <param name="address1">The first <see cref="Address"/> instance to compare.</param>
        /// <param name="address2">The second <see cref="Address"/> instance to compare.</param>
        /// <returns><see langword="true"/> if the specified <see cref="Address"/> instances are equal; otherwise, <see
        /// langword="false"/>.</returns>
        public static bool operator ==(Address address1, Address address2)
        {
            return address1.Equals(address2);
        }

        /// <summary>
        /// Determines whether two <see cref="Address"/> instances are not equal.
        /// </summary>
        /// \remark <remarks>This operator uses the <see cref="Address.Equals(Address)"/> method to determine
        /// equality.</remarks>
        /// <param name="address1">The first <see cref="Address"/> instance to compare.</param>
        /// <param name="address2">The second <see cref="Address"/> instance to compare.</param>
        /// <returns><see langword="true"/> if the two <see cref="Address"/> instances are not equal; otherwise, <see langword="false"/>.</returns>
        public static bool operator !=(Address address1, Address address2)
        {
            return !address1.Equals(address2);
        }

        /// <summary>
        /// Compares two addresses using the column and row numbers
        /// </summary>
        /// <param name="other"> Other address</param>
        /// <returns>-1 if the other address is greater, 0 if equal and 1 if smaller</returns>
        public int CompareTo(Address other)
        {
            long thisCoordinate = (long)Column * (long)Worksheet.MaxRowNumber + Row;
            long otherCoordinate = (long)other.Column * (long)Worksheet.MaxRowNumber + other.Row;
            return thisCoordinate.CompareTo(otherCoordinate);
        }

        /// <summary>
        /// Determines whether one specified <see cref="Address"/> is less/smaller than another specified <see cref="Address"/>.
        /// </summary>
        /// <param name="left">Left address</param>
        /// <param name="right">Right address</param>
        /// <returns>True, if the left address is less/smaller than the right one</returns>
        public static bool operator <(Address left, Address right)
        {
            return left.CompareTo(right) < 0;
        }

        /// <summary>
        /// Determines whether one specified <see cref="Address"/> is less/smaller or equal than another specified <see cref="Address"/>.
        /// </summary>
        /// <param name="left">Left address</param>
        /// <param name="right">Right address</param>
        /// <returns>True, if the left address is less/smaller than, or equal to the right one</returns>
        public static bool operator <=(Address left, Address right)
        {
            return left.CompareTo(right) <= 0;
        }

        /// <summary>
        /// Determines whether one specified <see cref="Address"/> is greater/larger than another specified <see cref="Address"/>.
        /// </summary>
        /// <param name="left">Left address</param>
        /// <param name="right">Right address</param>
        /// <returns>True, if the left address is greater/larger than the right one</returns>
        public static bool operator >(Address left, Address right)
        {
            return left.CompareTo(right) > 0;
        }

        /// <summary>
        /// Determines whether one specified <see cref="Address"/> is greater/larger or equal than another specified <see cref="Address"/>.
        /// </summary>
        /// <param name="left">Left address</param>
        /// <param name="right">Right address</param>
        /// <returns>True, if the left address is greater/larger than, or equal to the right one</returns>
        public static bool operator >=(Address left, Address right)
        {
            return left.CompareTo(right) >= 0;
        }

        /// <summary>
        /// Explicit conversion from string to Address
        /// </summary>
        /// <param name="address">Address expression</param>
        public static explicit operator Address(string address)
        {
            return new Address(address);
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
