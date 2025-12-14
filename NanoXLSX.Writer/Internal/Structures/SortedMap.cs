/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way  
 * Copyright Raphael Stoeckli © 2025
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using NanoXLSX.Interfaces;
using System.Collections.Generic;

namespace NanoXLSX.Internal
{
    /// <summary>
    /// Class to manage key value pairs (string / string). The entries are in the order how they were added
    /// </summary>
    /// \remark <remarks>This class is only for internal use. Use the high level API (e.g. class Workbook) to manipulate data and create Excel files</remarks>
    internal class SortedMap : ISortedMap
    {
        private readonly List<string> indexEntries;
        private readonly Dictionary<IFormattableText, int> index;
        private List<IFormattableText> keys;

        /// <summary>
        /// Number of map entries
        /// </summary>
        public int Count { get; private set; }

        /// <summary>
        /// Gets the keys of the map as list
        /// </summary>
        public IEnumerable<IFormattableText> Keys => keys;

        /// <summary>
        /// Gets the number of map entries (interface implementation)
        /// </summary>
        int ISortedMap.Count => keys.Count;

        /// <summary>
        /// Default constructor
        /// </summary>
        public SortedMap()
        {
            keys = new List<IFormattableText>();
            indexEntries = new List<string>();
            index = new Dictionary<IFormattableText, int>();
            Count = 0;
        }

        /// <summary>
        /// Method to add a key value pair (IXmlFormattableText as key and its index in the worksheet as value)
        /// </summary>
        /// <param name="text">Text (Key) as string</param>
        /// <param name="referenceIndex">Reference index as string</param>
        /// <returns>Returns the resolved string (either added or returned from an existing entry) of the reference index</returns>
        public string Add(IFormattableText text, string referenceIndex)
        {
            if (index.ContainsKey(text))
            {
                return indexEntries[index[text]];
            }
            index.Add(text, Count);
            Count++;
            keys.Add(text);
            indexEntries.Add(referenceIndex);
            return referenceIndex;
        }

        /// <summary>
        /// Method to add a key value pair (IFormattableText as key and its index in the worksheet as value)
        /// </summary>
        /// <param name="text">Text (Key) as string</param>
        /// <param name="referenceIndex">Reference index as string</param>
        /// <returns>Returns the resolved string (either added or returned from an existing entry) of the reference index</returns>
        string ISortedMap.Add(IFormattableText text, string referenceIndex)
        {
            return Add(text, referenceIndex);
        }
    }
}
