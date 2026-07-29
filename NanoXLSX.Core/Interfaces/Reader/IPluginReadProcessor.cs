/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System;

namespace NanoXLSX.Interfaces.Reader
{
    /// <summary>
    /// Interface, used by base reader plug-ins that are not handling a stream, but data in <see cref="IPluginReader.Workbook"/>
    /// </summary>
    internal interface IPluginReadProcessor : IPluginReader
    {

        /// <summary>
        /// Optional reader options
        /// </summary>
        IOptions Options { get; set; }

        /// <summary>
        /// Reference to a handler of in-line plugins, to be used for post operations in the <see cref="IPlugin.Execute"/> method
        /// </summary>
        Action<Workbook, string, IOptions, int?> InlinePluginHandler { get; set; }

        /// <summary>
        /// Initialization method
        /// </summary>
        /// <param name="workbook">Workbook instance where read data is placed</param>
        /// <param name="readerOptions">Optional reader options</param>
        /// <param name="inlinePluginHandler">Reference to the a handler action, to be used for post operations in processor methods</param>
        void Init(Workbook workbook, IOptions readerOptions, Action<Workbook, string, IOptions, int?> inlinePluginHandler);

    }
}
