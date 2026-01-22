/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System;
using System.Collections.Generic;

namespace NanoXLSX.Interfaces.Reader
{
    /// <summary>
    /// Interface, used by worksheet readers
    /// </summary>
    internal interface IWorksheetReader : IPluginBaseReader
    {
        /// <summary>
        /// Gets or sets the (r)ID (1-based) of the currently processed worksheet.
        /// </summary>
        int CurrentWorksheetID { get; set; }

        /// <summary>
        /// Gets or Sets the list of the shared strings. The index of the list corresponds to the index, defined in cell values
        /// </summary>
        List<string> SharedStrings { get; set; }
    }
}
