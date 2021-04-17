/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way  
 * Copyright Raphael Stoeckli © 2021
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System.Collections.Generic;

namespace NanoXLSX
{
    /// <summary>
    /// Struct representing a cell range with a start and end address
    /// </summary>
    public struct Range
    {
        /// <summary>
        /// End address of the range
        /// </summary>
        public Address EndAddress;
        /// <summary>
        /// Start address of the range
        /// </summary>
        public Address StartAddress;

        /// <summary>
        /// Constructor with addresses as arguments
        /// </summary>
        /// <param name="start">Start address of the range</param>
        /// <param name="end">End address of the range</param>
        public Range(Address start, Address end)
        {
            StartAddress = start;
            EndAddress = end;
        }

        /// <summary>
        /// Constructor with a range string as argument
        /// </summary>
        /// <param name="range">Address range (e.g. 'A1:B12')</param>
        public Range(string range)
        {
            Range r = Cell.ResolveCellRange(range);
            StartAddress = r.StartAddress;
            EndAddress = r.EndAddress;
        }

        /// <summary>
        /// Gets a list of all addresses between the start and end address
        /// </summary>
        /// <returns>List of Addresses</returns>
        public IReadOnlyList<Address> ResolveEnclosedAddresses()
        {
            int startColumn, endColumn, startRow, endRow;
            if (StartAddress.Column <= EndAddress.Column)
            {
                startColumn = this.StartAddress.Column;
                endColumn = this.EndAddress.Column;
            }
            else
            {
                endColumn = this.StartAddress.Column;
                startColumn = this.EndAddress.Column;
            }
            if (StartAddress.Row <= EndAddress.Row)
            {
                startRow = this.StartAddress.Row;
                endRow = this.EndAddress.Row;
            }
            else
            {
                endRow = this.StartAddress.Row;
                startRow = this.EndAddress.Row;
            }
            List<Address> addresses = new List<Address>();
            for(int r = startRow; r <= endRow; r++)
            {
                for(int c = startColumn; c <= endColumn; c++)
                {
                    addresses.Add(new Address(c, r));
                }
            }
            return addresses;
        }

        /// <summary>
        /// Overwritten ToString method
        /// </summary>
        /// <returns>Returns the range (e.g. 'A1:B12')</returns>
        public override string ToString()
        {
            return StartAddress + ":" + EndAddress;
        }

    }

}
