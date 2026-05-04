/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System;
using System.Collections;
using System.Collections.Generic;

namespace NanoXLSX.Internal
{
    /// <summary>
    /// Non-materializing read-only view over the internal cell dictionary that exposes cells keyed by
    /// their rendered address string (e.g. "A1"). All read paths delegate directly to the backing
    /// dictionary, translating keys on-the-fly without allocating a snapshot copy.
    /// Mutation is intentionally unsupported — use <c>Worksheet.AddCell</c> / <c>Worksheet.RemoveCell</c>.
    /// </summary>
    internal sealed class StringKeyedCellView : IReadOnlyDictionary<string, Cell>
    {
        private readonly Dictionary<CellKey, Cell> store;

        internal StringKeyedCellView(Dictionary<CellKey, Cell> store)
        {
            this.store = store;
        }

        /// <inheritdoc/>
        public Cell this[string key]
        {
            get
            {
                Cell.ResolveCellCoordinate(key, out int col, out int row);
                return store[new CellKey(col, row)];
            }
        }

        /// <inheritdoc/>
        public int Count
        {
            get { return store.Count; }
        }

        /// <inheritdoc/>
        public IEnumerable<string> Keys
        {
            get
            {
                foreach (CellKey k in store.Keys)
                {
                    yield return Cell.ResolveCellAddress(k.Column, k.Row);
                }
            }
        }

        /// <inheritdoc/>
        public IEnumerable<Cell> Values
        {
            get { return store.Values; }
        }

        /// <inheritdoc/>
        public bool ContainsKey(string key)
        {
            if (string.IsNullOrEmpty(key))
            {
                return false;
            }
            try
            {
                Cell.ResolveCellCoordinate(key, out int col, out int row);
                return store.ContainsKey(new CellKey(col, row));
            }
            catch (Exception)
            {
                return false;
            }
        }

        /// <inheritdoc/>
        public bool TryGetValue(string key, out Cell value)
        {
            if (string.IsNullOrEmpty(key))
            {
                value = null;
                return false;
            }
            try
            {
                Cell.ResolveCellCoordinate(key, out int col, out int row);
                return store.TryGetValue(new CellKey(col, row), out value);
            }
            catch (Exception)
            {
                value = null;
                return false;
            }
        }

        /// <inheritdoc/>
        public IEnumerator<KeyValuePair<string, Cell>> GetEnumerator()
        {
            foreach (KeyValuePair<CellKey, Cell> kv in store)
            {
                yield return new KeyValuePair<string, Cell>(
                    Cell.ResolveCellAddress(kv.Key.Column, kv.Key.Row),
                    kv.Value);
            }
        }

        /// <inheritdoc/>
        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }
    }
}
