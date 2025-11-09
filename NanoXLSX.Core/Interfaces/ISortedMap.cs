/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2025
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System.Collections.Generic;

namespace NanoXLSX.Interfaces
{
    /// <summary>
    /// Interface to represent a sorted map with IFormattableText as key and string as value
    /// </summary>
    public interface ISortedMap
    {
        /// <summary>
        /// Number of map entries
        /// </summary>
        int Count { get; }

        /// <summary>
        /// Gets the keys of the map as list
        /// </summary>
        List<IFormattableText> Keys { get; }

        /// <summary>
        /// Method to add a key value pair (IFormattableText as key and its index in the worksheet as value)
        /// </summary>
        /// <param name="text">Text (Key) as string</param>
        /// <param name="referenceIndex">Reference index as string</param>
        /// <returns>Returns the resolved string (either added or returned from an existing entry) of the reference index</returns>
        string Add(IFormattableText text, string referenceIndex);

    }
}
