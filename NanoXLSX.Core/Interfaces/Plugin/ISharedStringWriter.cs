/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2025
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

namespace NanoXLSX.Interfaces.Writer
{
    /// <summary>
    /// Interface, used by shared string writers
    /// </summary>
    internal interface ISharedStringWriter : IPluginWriter
    {
        /// <summary>
        /// Sorted map that contains the shared strings
        /// </summary>
        ISortedMap SharedStrings { get; }

        /// <summary>
        /// Total number of shared strings
        /// </summary>
        int SharedStringsTotalCount { get; set; }
    }
}
