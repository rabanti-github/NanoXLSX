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
        /// Constructor with addresses as arguments. The addresses are automatically swapped if the start address is greater than the end address
        /// </summary>
        /// <param name="start">Start address of the range</param>
        /// <param name="end">End address of the range</param>
        public Range(Address start, Address end)
        {
            if (start.CompareTo(end) < 0)
            {
                StartAddress = start;
                EndAddress = end;
            }
            else
            {
                StartAddress = end;
                EndAddress = start;
            }
        }

        /// <summary>
        /// Constructor with a range string as argument. The addresses are automatically swapped if the start address is greater than the end address
        /// </summary>
        /// <param name="range">Address range (e.g. 'A1:B12')</param>
        public Range(string range)
        {
            Range r = Cell.ResolveCellRange(range);
            if (r.StartAddress.CompareTo(r.EndAddress) < 0)
            {
                StartAddress = r.StartAddress;
                EndAddress = r.EndAddress;
            }
            else
            {
                StartAddress = r.EndAddress;
                EndAddress = r.StartAddress;
            }
        }

        /// <summary>
        /// Gets a list of all addresses between the start and end address
        /// </summary>
        /// <returns>List of Addresses</returns>
        public IReadOnlyList<Address> ResolveEnclosedAddresses()
        {
            IEnumerable<Address> range = Cell.GetCellRange(this.StartAddress, this.EndAddress);
            return new List<Address>(range);
        }

        /// <summary>
        /// Overwritten ToString method
        /// </summary>
        /// <returns>Returns the range (e.g. 'A1:B12')</returns>
        public override string ToString()
        {
            return StartAddress + ":" + EndAddress;
        }

        public override bool Equals(object obj)
        {
            if (obj == null || !(obj is Range))
            {
                return false;
            }
            Range other = (Range)obj;
            return this.StartAddress.Equals(other.StartAddress) && this.EndAddress.Equals(other.EndAddress);
        }

    }

}
