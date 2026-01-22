/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2025
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System;
using System.IO;

namespace NanoXLSX.Interfaces.Reader
{
    /// <summary>
    /// Interface, used by in-line (queue) plug-ins in XML reader classes 
    /// </summary>
    internal interface IPluginInlineReader : IPluginReader
    {
        /// <summary>
        /// Reference to the a handler action, to be used for post operations in reader methods. Only relevant for in-line plug-ins, therefore null for queue plug-ins.
        /// </summary>
        Action<MemoryStream, Workbook, string, IOptions, int?> InlinePluginHandler { get; set; }

        /// <summary>
        /// Initialization method
        /// </summary>
        /// <param name="stream">Stream, containing the XML file to red</param>
        /// <param name="workbook">Workbook instance where read data is placed</param>
        /// <param name="readerOptions">Optional reader options</param>
        /// <param name="index">Optional index, e.g. for worksheet identification</param>
        void Init(MemoryStream stream, Workbook workbook, IOptions readerOptions, int? index = null);

    }
}
