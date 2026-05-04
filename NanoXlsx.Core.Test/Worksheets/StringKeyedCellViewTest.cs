using System.Collections;
using System.Collections.Generic;
using NanoXLSX.Internal;
using Xunit;

namespace NanoXLSX.Test.Core.InternalTest
{
    /// <summary>
    /// Tests for <see cref="StringKeyedCellView"/> edge cases that are not covered by the
    /// higher-level <see cref="CellValuesTest"/>: null/empty key guards and the invalid-address
    /// exception path in both ContainsKey and TryGetValue.
    /// </summary>
    public class StringKeyedCellViewTest
    {
        [Fact(DisplayName = "ContainsKey: null key returns false without throwing")]
        public void ContainsKey_NullKey_ReturnsFalse()
        {
            StringKeyedCellView view = BuildView((0, 0, "x"));
            Assert.False(view.ContainsKey(null));
        }

        [Theory(DisplayName = "ContainsKey: empty/mull string key returns false without throwing")]
        [InlineData("")]
        [InlineData(null)]
        public void ContainsKey_EmptyKey_ReturnsFalse(string key)
        {
            StringKeyedCellView view = BuildView((0, 0, "x"));
            Assert.False(view.ContainsKey(key));
        }

        [Theory(DisplayName = "ContainsKey: malformed address returns false without throwing")]
        [InlineData("123")]       // digits only — no column letters
        [InlineData("!!!")]       // non-alphanumeric
        [InlineData("AAAAA1")]    // column part too long (> 3 letters)
        [InlineData("A0")]        // row 0 is out of range (1-based in Excel)
        [InlineData("A")]         // no row number
        [InlineData(" A7")]       // leading space
        public void ContainsKey_InvalidAddress_ReturnsFalse(string key)
        {
            StringKeyedCellView view = BuildView((0, 0, "x"));
            Assert.False(view.ContainsKey(key));
        }

        [Theory(DisplayName = "TryGetValue: empty/null key returns false and null cell")]
        [InlineData("")]
        [InlineData(null)]
        public void TryGetValue_NullKey_ReturnsFalse(string key)
        {
            StringKeyedCellView view = BuildView((0, 0, "x"));
            bool found = view.TryGetValue(key, out Cell cell);
            Assert.False(found);
            Assert.Null(cell);
        }

        [Theory(DisplayName = "TryGetValue: malformed address returns false and null cell")]
        [InlineData("123")]
        [InlineData("!!!")]
        [InlineData("AAAAA1")]
        [InlineData("A0")]
        [InlineData("A")]
        [InlineData("ZZZZZ9999")]
        [InlineData(" A7")]
        public void TryGetValue_InvalidAddress_ReturnsFalseAndNullCell(string key)
        {
            StringKeyedCellView view = BuildView((0, 0, "x"));
            bool found = view.TryGetValue(key, out Cell cell);
            Assert.False(found);
            Assert.Null(cell);
        }

        [Fact(DisplayName = "TryGetValue: valid address for existing cell returns true and correct cell")]
        public void TryGetValue_ExistingKey_ReturnsTrueAndCell()
        {
            StringKeyedCellView view = BuildView((2, 4, "hello"));
            bool found = view.TryGetValue("C5", out Cell cell);
            Assert.True(found);
            Assert.NotNull(cell);
            Assert.Equal("hello", cell.Value);
        }

        [Fact(DisplayName = "TryGetValue: valid address for absent cell returns false and null cell")]
        public void TryGetValue_AbsentKey_ReturnsFalseAndNullCell()
        {
            StringKeyedCellView view = BuildView((0, 0, "x"));
            bool found = view.TryGetValue("Z99", out Cell cell);
            Assert.False(found);
            Assert.Null(cell);
        }

        [Fact(DisplayName = "IEnumerable.GetEnumerator: non-generic enumerator yields all entries")]
        public void IEnumerable_GetEnumerator_YieldsAllEntries()
        {
            StringKeyedCellView view = BuildView((2, 4, "hello"), (0, 0, "world"));
            IEnumerable enumerable = view;
            int count = 0;
            foreach (object item in enumerable)
            {
                Assert.IsType<KeyValuePair<string, Cell>>(item);
                count++;
            }
            Assert.Equal(2, count);
        }

        private static StringKeyedCellView BuildView(params (int col, int row, object value)[] entries)
        {
            Dictionary<CellKey, Cell> store = new Dictionary<CellKey, Cell>();
            foreach (var (col, row, value) in entries)
            {
                store[new CellKey(col, row)] = new Cell(value, Cell.CellType.Default, col, row);
            }
            return new StringKeyedCellView(store);
        }

    }
}
