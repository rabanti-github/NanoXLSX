/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way  
 * Copyright Raphael Stoeckli © 2025
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System;
using System.Collections.Generic;
using System.Linq;

namespace NanoXLSX
{
    /// <summary>
    /// Struct representing a cell range with a start and end address
    /// </summary>
    public struct Range
    {
        private readonly Address startAddress;
        private readonly Address endAddress;

        /// <summary>
        /// End address of the range
        /// </summary>
        public Address EndAddress { get => endAddress; }
        /// <summary>
        /// Start address of the range
        /// </summary>
        public Address StartAddress { get => startAddress; }

        /// <summary>
        /// Constructor with addresses as arguments. The addresses are automatically swapped if the start address is greater than the end address.
        /// Referencing modifiers ($) for rows and columns can be passed through the address type of the address objects
        /// </summary>
        /// <param name="start">Start address of the range</param>
        /// <param name="end">End address of the range</param>
        public Range(Address start, Address end)
        {
            if (start.CompareTo(end) < 0)
            {
                this.startAddress = start;
                this.endAddress = end;
            }
            else
            {
                this.startAddress = end;
                this.endAddress = start;
            }
        }

        /// <summary>
        /// Constructor with start and end rows and columns as arguments. The addresses are automatically swapped if the start address is greater than the end address.
        /// Referencing modifiers ($) for rows and columns are not considered
        /// </summary>
        /// <param name="startColumn">Start column number (zero based) of the range</param>
        /// <param name="startRow">Start row number (zero based) of the range</param>
        /// <param name="endColumn">End column number (zero based) of the range</param>
        /// <param name="endRow">End row number (zero based) of the range</param>
        public Range(int startColumn, int startRow, int endColumn, int endRow) : this(new Address(startColumn, startRow), new Address(endColumn, endRow))
        {
        }


        /// <summary>
        /// Constructor with a range string as argument. The addresses are automatically swapped if the start address is greater than the end address.
        /// Referencing modifiers ($) for rows and columns can be defined in the passed string
        /// </summary>
        /// <param name="range">Address range (e.g. 'A1:B12')</param>
        public Range(string range)
        {
            Range r = Cell.ResolveCellRange(range);
            if (r.StartAddress.CompareTo(r.EndAddress) < 0)
            {
                this.startAddress = r.StartAddress;
                this.endAddress = r.EndAddress;
            }
            else
            {
                this.startAddress = r.EndAddress;
                this.endAddress = r.StartAddress;
            }
        }

        /// <summary>
        /// Gets whether another range is completely enclosed by this range
        /// </summary>
        /// <param name="other">Other range to check</param>
        /// <returns>True if the other range is completely enclosed. False if only partial overlapping or not intersecting</returns>
        public bool Contains(Range other)
        {
            return this.StartAddress.Column <= other.StartAddress.Column &&
                   this.EndAddress.Column >= other.EndAddress.Column &&
                   this.StartAddress.Row <= other.StartAddress.Row &&
                   this.EndAddress.Row >= other.EndAddress.Row;
        }

        /// <summary>
        /// Determines whether an address is within this range
        /// </summary>
        /// <param name="address">Address to check</param>
        /// <returns>True if the address is part of this range, otherwise false</returns>
        public bool Contains(Address address)
        {
            return address.Column >= this.startAddress.Column &&
                address.Column <= this.endAddress.Column &&
                address.Row >= this.startAddress.Row &&
                address.Row <= this.endAddress.Row;
        }

        /// <summary>
        /// Determines whether the passed range overlaps with this range
        /// </summary>
        /// <param name="other">Range to check for overlapping</param>
        /// <returns>True if overlapping, otherwise false</returns>
        public bool Overlaps(Range other)
        {
            return !(this.EndAddress.Row < other.StartAddress.Row || this.StartAddress.Row > other.EndAddress.Row ||
                     this.EndAddress.Column < other.StartAddress.Column || this.StartAddress.Column > other.EndAddress.Column);
        }


        /// <summary>
        /// Gets a list of all addresses between the start and end address
        /// </summary>
        /// <returns>List of Addresses</returns>
        /// \remark <remarks>Use this function with caution. Very big ranges may result to hundred of Millions or even Billions of cells. This may lead to an extremely high memory consumptions or even a crash of the application</remarks>
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

        // Operator overloads
        /// <summary>
        /// Compares two objects whether they are ranges and equal. The cell types (possible $ prefix) are considered. This method reflects <see cref="Equals(object)"/>
        /// </summary>
        /// <param name="range1">First range object</param>
        /// <param name="range2">Second range object</param>
        /// <returns>True, if both objects are equal, otherwise false</returns>
        public static bool operator == (Range range1, Range range2)
        {
            return range1.Equals(range2);
        }

        /// <summary>
        /// Compares two objects whether they not equal. This method reflects the inverted method of <see cref="Equals(object)"/>
        /// </summary>
        /// <param name="range1">First range object</param>
        /// <param name="range2">Second range object</param>
        /// <returns>False, if both objects are equal, otherwise true</returns>
        public static bool operator != (Range range1, Range range2)
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
