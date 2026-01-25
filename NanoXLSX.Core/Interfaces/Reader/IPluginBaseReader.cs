/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System;
using System.IO;

namespace NanoXLSX.Interfaces.Reader
{
    /// <summary>
    /// Interface, used by base plug-ins in XML reader classes 
    /// </summary>
    internal interface IPluginBaseReader : IPluginReader
    {

        /// <summary>
        /// Optional reader options
        /// </summary>
        IOptions Options { get; set; }

        /// <summary>
        /// Initialization method
        /// </summary>
        /// <param name="stream">Stream, containing the XML file to red</param>
        /// <param name="workbook">Workbook instance where read data is placed</param>
        /// <param name="readerOptions">Optional reader options</param>
        /// <param name="inlinePluginHandler">Reference to the a handler action, to be used for post operations in reader methods</param>
        void Init(MemoryStream stream, Workbook workbook, IOptions readerOptions, Action<MemoryStream, Workbook, string, IOptions, int?> inlinePluginHandler);

        /// <summary>
        /// Reference to a handler of in-line plugins, to be used for post operations in the <see cref="IPlugin.Execute"/> method
        /// </summary>
        Action<MemoryStream, Workbook, string, IOptions, int?> InlinePluginHandler { get; set; }

    }
}
