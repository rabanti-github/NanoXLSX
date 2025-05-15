/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2025
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System.Collections.Generic;
using NanoXLSX.Interfaces.Reader;

namespace NanoXLSX.Interfaces.Plugin
{
    /// <summary>
    /// Interface, used by shared string readers
    /// </summary>
    internal interface ISharedStringReader : IPlugInReader
    {
        /// <summary>
        /// Resolved list of shared strings.
        /// The indices of the shared strings are defined by the order of the strings in the list.
        /// </summary>
        List<string> SharedStrings { get; }
    }
}
