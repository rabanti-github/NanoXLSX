/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way  
 * Copyright Raphael Stoeckli © 2018
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

namespace NanoXLSX
{
    public partial class Cell
    {
        /// <summary>
        /// Struct representing the cell address as column and row (zero based)
        /// </summary>
        public struct Address
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
            public AddressType Type;

            /// <summary>
            /// Constructor with row and column as arguments
            /// </summary>
            /// <param name="column">Column number (zero based)</param>
            /// <param name="row">Row number (zero based)</param>
            /// <param name="type">Optional referencing type of the address</param>
            public Address(int column, int row, AddressType type = AddressType.Default)
            {
                Column = column;
                Row = row;
                Type = type;
            }

            /// <summary>
            /// Constructor with address as string
            /// </summary>
            /// <param name="address">Address string (e.g. 'A1:B12')</param>
            /// <param name="type">Optional referencing type of the address</param>
            public Address(string address, AddressType type = AddressType.Default)
            {
                Type = type;
                ResolveCellCoordinate(address, out Column, out Row);
            }

            /// <summary>
            /// Returns the combined Address
            /// </summary>
            /// <returns>Address as string in the format A1 - XFD1048576</returns>
            public string GetAddress()
            {
                return ResolveCellAddress(Column, Row, Type);
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
                if (Row == o.Row && Column == o.Column) { return true; }
                else { return false; }
            }

        }
    }
}