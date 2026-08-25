/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System;

namespace NanoXLSX.Internal
{
    /// <summary>
    /// Compact, hash-efficient key for the internal cell dictionary. Uses integer (col, row) coordinates
    /// instead of rendered address strings, eliminating string allocation on every cell insert/lookup.
    /// </summary>
    internal readonly struct CellKey : IEquatable<CellKey>
    {
        internal readonly int Column;
        internal readonly int Row;

        /// <summary>
        /// COnstructor with all parameters
        /// </summary>
        /// <param name="column">Column number</param>
        /// <param name="row">Row number</param>
        internal CellKey(int column, int row)
        {
            Column = column;
            Row = row;
        }

        /// <summary>
        /// Returns whether two instances of CellKey are the same
        /// </summary>
        /// <param name="other">The reference CellKey with which to compare</param>
        /// <returns>True if this instance and the other are the same</returns>
        public bool Equals(CellKey other)
        {
            return Column == other.Column && Row == other.Row;
        }

        /// <summary>
        /// Returns whether two instances are the same
        /// </summary>
        /// <param name="obj">The reference object with which to compare</param>
        /// <returns>True if this instance and the other are the same</returns>
        public override bool Equals(object obj)
        {
            return obj is CellKey other && Equals(other);
        }

        // Excel max: 16 384 columns (14 bits) × 1 048 576 rows (20 bits) — fits cleanly in 34 bits,
        // so a simple multiply+XOR gives a collision-free hash across the valid address space.
        /// <summary>
        /// Gets the hash code of the cell key
        /// </summary>
        /// <returns>Hash code</returns>
        public override int GetHashCode()
        {
            return (Row * 16384) ^ Column;
        }

        /// <summary>
        /// returns the String representation of the cell key
        /// </summary>
        /// <returns>Cell address</returns>
        public override string ToString()
        {
            return Cell.ResolveCellAddress(Column, Row);
        }
    }
}
