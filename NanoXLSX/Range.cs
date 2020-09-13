/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way  
 * Copyright Raphael Stoeckli © 2020
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

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
        /// Overwritten ToString method
        /// </summary>
        /// <returns>Returns the range (e.g. 'A1:B12')</returns>
        public override string ToString()
        {
            return StartAddress + ":" + EndAddress;
        }

    }

}
