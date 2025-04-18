/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2025
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System.IO;
using NanoXLSX.Interfaces.Plugin;

namespace NanoXLSX.Interfaces.Reader
{
    /// <summary>
    /// Interface, used by XML reader classes 
    /// </summary>
    internal interface IPlugInReader : IPlugIn
    {
        /// <summary>
        /// Gets or replaces the workbook instance, defined by the constructor
        /// </summary>
        Workbook Workbook { get; set; }

        /// <summary>
        /// Initialization method
        /// </summary>
        /// <param name="stream">Stream, containing the XML file to red</param>
        /// <param name="workbook">Workbook instance where read data is placed</param>
        /// <param name="readerOptions">Optional reader options</param>
        void Init(MemoryStream stream, Workbook workbook, IOptions readerOptions);

    }
}
