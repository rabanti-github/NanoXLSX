/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

namespace NanoXLSX.Interfaces.Reader
{
    /// <summary>
    /// Interface, used by in-line (queue) reader plug-ins that are not handling a stream, but data in <see cref="IPluginReader.Workbook"/>
    /// </summary>
    internal interface IPluginInlineReadProcessor : IPluginReader
    {
        /// <summary>
        /// Initialization method
        /// </summary>
        /// <param name="workbook">Workbook instance where read data is placed</param>
        /// <param name="readerOptions">Optional reader options</param>
        /// <param name="index">Optional index, e.g. for worksheet identification</param>
        void Init(Workbook workbook, IOptions readerOptions, int? index = null);

    }
}
