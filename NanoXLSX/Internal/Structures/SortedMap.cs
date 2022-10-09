/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way  
 * Copyright Raphael Stoeckli © 2022
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using NanoXLSX.Shared.Interfaces;
using System.Collections.Generic;

namespace NanoXLSX.Internal
{
    /// <summary>
    /// Class to manage key value pairs (string / string). The entries are in the order how they were added
    /// </summary>
    public class SortedMap
    {
        private int count;
        private readonly List<IFormattableText> valueEntries;
        private readonly List<string> indexEntries;
        private readonly Dictionary<IFormattableText, int> index;

        /// <summary>
        /// Number of map entries
        /// </summary>
        public int Count
        {
            get { return count; }
        }

        /// <summary>
        /// Gets the keys of the map as list
        /// </summary>
        public List<IFormattableText> Keys
        {
            get { return valueEntries; }
        }

        /// <summary>
        /// Default constructor
        /// </summary>
        public SortedMap()
        {
            valueEntries = new List<IFormattableText>();
            indexEntries = new List<string>();
            index = new Dictionary<IFormattableText, int>();
            count = 0;
        }

        /// <summary>
        /// Method to add a key value pair (IFormattableText as key and its index in the worksheet as value)
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
            index.Add(text, count);
            count++;
            valueEntries.Add(text);
            indexEntries.Add(referenceIndex);
            return referenceIndex;
        }
    }
}
