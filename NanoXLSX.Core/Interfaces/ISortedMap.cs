using System;
using System.Collections.Generic;
using System.Text;

namespace NanoXLSX.Interfaces
{
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
