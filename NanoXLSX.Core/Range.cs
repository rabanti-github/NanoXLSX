/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way  
 * Copyright Raphael Stoeckli © 2025
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

        /// <summary>
        /// Compares two objects whether they are ranges and equal. The cell types (possible $ prefix) are considered 
        /// </summary>
        /// <param name="obj">Other object to compare</param>
        /// <returns>True if the two objects are the same range</returns>
        public override bool Equals(object obj)
        {
            if (!(obj is Range))
            {
                return false;
            }
            Range other = (Range)obj;
            return this.StartAddress.Equals(other.StartAddress) && this.EndAddress.Equals(other.EndAddress);
        }

        /// <summary>
        /// Gets the hash code of the range object according to its string representation
        /// </summary>
        /// <returns>Hash code of the range</returns>
        public override int GetHashCode()
        {
            return this.ToString().GetHashCode();
        }

        // Operator overloads
        public static bool operator ==(Range range1, Range range2)
        {
            return range1.Equals(range2);
        }

        public static bool operator !=(Range range1, Range range2)
        {
            return !range1.Equals(range2);
        }


        /// <summary>
        /// Creates a (dereferenced, if applicable) deep copy of this range
        /// </summary>
        /// <returns>Copy of this range</returns>
        internal Range Copy()
        {
            return new Range(this.StartAddress.Copy(), this.EndAddress.Copy());
        }

    }

}
