/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way  
 * Copyright Raphael Stoeckli © 2018
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
            private List<string> keyEntries;
            private List<string> valueEntries;
            private Dictionary<string, int> index;

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
            /// Gets the values of the map as values
            /// </summary>
            public List<string> Values
            {
                get { return valueEntries; }
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
            /// Indexer to get the specific value by the key
            /// </summary>
            /// <param name="key">Key to corresponding value. Returns null if not found</param>
            public string this[string key]
            {
                get
                {
                    if (index.ContainsKey(key))
                    {
                        return valueEntries[index[key]];
                    }
                    return null;
                }
            }

            /// <summary>
            /// Adds a key value pair to the map. If the key already exists, only its index will be returned
            /// </summary>
            /// <param name="key">Key of the tuple</param>
            /// <param name="value">Value of the tuple</param>
            /// <returns>Position of the tuple in the map as index (zero-based)</returns>
            public int Add(string key, string value)
            {
                if (index.ContainsKey(key))
                {
                    return index[key];
                }

                index.Add(key, count);
                keyEntries.Add(key);
                valueEntries.Add(value);
                count++;
                return count - 1;
            }

            /// <summary>
            /// Gets whether the specified key exists in the map
            /// </summary>
            /// <param name="key">Key to check</param>
            /// <returns>True if the entry exists, otherwise false</returns>
            public bool ContainsKey(string key)
            {
                return index.ContainsKey(key);
            }

        }
}