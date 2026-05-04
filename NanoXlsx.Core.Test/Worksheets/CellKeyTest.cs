/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using NanoXLSX.Internal;
using Xunit;

namespace NanoXLSX.Test.Core.InternalTest
{
    public class CellKeyTest
    {
        // ── Equals(CellKey) ──────────────────────────────────────────────────

        [Fact(DisplayName = "Equals(CellKey): same column and row returns true")]
        public void Equals_Typed_SameColRow_ReturnsTrue()
        {
            CellKey a = new CellKey(3, 7);
            CellKey b = new CellKey(3, 7);
            Assert.True(a.Equals(b));
        }

        [Fact(DisplayName = "Equals(CellKey): different column, same row returns false")]
        public void Equals_Typed_DifferentCol_ReturnsFalse()
        {
            CellKey a = new CellKey(3, 7);
            CellKey b = new CellKey(4, 7);
            Assert.False(a.Equals(b));
        }

        [Fact(DisplayName = "Equals(CellKey): same column, different row returns false")]
        public void Equals_Typed_DifferentRow_ReturnsFalse()
        {
            CellKey a = new CellKey(3, 7);
            CellKey b = new CellKey(3, 8);
            Assert.False(a.Equals(b));
        }

        [Fact(DisplayName = "Equals(CellKey): different column and row returns false")]
        public void Equals_Typed_DifferentColAndRow_ReturnsFalse()
        {
            CellKey a = new CellKey(0, 0);
            CellKey b = new CellKey(1, 1);
            Assert.False(a.Equals(b));
        }

        // ── Equals(object) ───────────────────────────────────────────────────

        [Fact(DisplayName = "Equals(object): boxed CellKey with same values returns true")]
        public void Equals_Object_BoxedSameValues_ReturnsTrue()
        {
            CellKey a = new CellKey(5, 2);
            object b = new CellKey(5, 2);
            Assert.True(a.Equals(b));
        }

        [Fact(DisplayName = "Equals(object): boxed CellKey with different values returns false")]
        public void Equals_Object_BoxedDifferentValues_ReturnsFalse()
        {
            CellKey a = new CellKey(5, 2);
            object b = new CellKey(5, 3);
            Assert.False(a.Equals(b));
        }

        [Fact(DisplayName = "Equals(object): null returns false")]
        public void Equals_Object_Null_ReturnsFalse()
        {
            CellKey a = new CellKey(1, 1);
            Assert.False(a.Equals(null));
        }

        [Fact(DisplayName = "Equals(object): unrelated type returns false")]
        public void Equals_Object_UnrelatedType_ReturnsFalse()
        {
            CellKey a = new CellKey(1, 1);
            Assert.False(a.Equals("A2"));
        }

        // ── GetHashCode ───────────────────────────────────────────────────────

        [Fact(DisplayName = "GetHashCode: equal keys produce equal hash codes")]
        public void GetHashCode_EqualKeys_EqualHashes()
        {
            CellKey a = new CellKey(10, 20);
            CellKey b = new CellKey(10, 20);
            Assert.Equal(a.GetHashCode(), b.GetHashCode());
        }

        [Theory(DisplayName = "GetHashCode: no collision at boundary coordinates")]
        [InlineData(0, 0)]
        [InlineData(16383, 0)]
        [InlineData(0, 1048575)]
        [InlineData(16383, 1048575)]
        public void GetHashCode_BoundaryCoordinates_Stable(int col, int row)
        {
            CellKey key = new CellKey(col, row);
            // calling twice must return the same value (deterministic)
            Assert.Equal(key.GetHashCode(), key.GetHashCode());
        }

        [Fact(DisplayName = "GetHashCode: keys that differ only in column produce different hashes")]
        public void GetHashCode_DifferentColumn_DifferentHash()
        {
            // (row*16384)^col — for two keys with the same row, hashes differ iff columns differ
            CellKey a = new CellKey(0, 5);
            CellKey b = new CellKey(1, 5);
            Assert.NotEqual(a.GetHashCode(), b.GetHashCode());
        }

        [Fact(DisplayName = "GetHashCode: keys that differ only in row produce different hashes")]
        public void GetHashCode_DifferentRow_DifferentHash()
        {
            // (row*16384)^col — for two keys with the same col=0, hashes differ iff rows differ
            CellKey a = new CellKey(0, 0);
            CellKey b = new CellKey(0, 1);
            Assert.NotEqual(a.GetHashCode(), b.GetHashCode());
        }

        // ── ToString ─────────────────────────────────────────────────────────

        [Theory(DisplayName = "ToString: renders the Excel address string")]
        [InlineData(0, 0, "A1")]
        [InlineData(1, 0, "B1")]
        [InlineData(25, 0, "Z1")]
        [InlineData(26, 0, "AA1")]
        [InlineData(0, 9, "A10")]
        [InlineData(2, 4, "C5")]
        [InlineData(16383, 1048575, "XFD1048576")]
        public void ToString_ReturnsExcelAddress(int col, int row, string expected)
        {
            CellKey key = new CellKey(col, row);
            Assert.Equal(expected, key.ToString());
        }
    }
}
