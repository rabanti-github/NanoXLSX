/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way  
 * Copyright Raphael Stoeckli © 2022
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System.Collections.Generic;

namespace NanoXLSX.LowLevel
{
    /// <summary>
    /// Class to manage key value pairs (string / string). The entries are in the order how they were added
    /// </summary>
    public class SortedMap
    {
        private int count;
        private readonly List<string> keyEntries;
        private readonly List<string> valueEntries;
        private readonly Dictionary<string, int> index;

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
        public List<string> Keys
        {
            get { return keyEntries; }
        }

        /// <summary>
        /// Default constructor
        /// </summary>
        public SortedMap()
        {
            keyEntries = new List<string>();
            valueEntries = new List<string>();
            index = new Dictionary<string, int>();
            count = 0;
        }

        /// <summary>
        /// Method to add a key value pair
        /// </summary>
        /// <param name="key">Key as string</param>
        /// <param name="value">Value as string</param>
        /// <returns>Returns the resolved string (either added or returned from an existing entry)</returns>
        public string Add(string key, string value)
        {
            if (index.ContainsKey(key))
            {
                return valueEntries[index[key]];
            }
            index.Add(key, count);
            count++;
            keyEntries.Add(key);
            valueEntries.Add(value);
            return value;
        }
    }
}
