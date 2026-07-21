/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System.IO.Compression;

namespace NanoXLSX.Interfaces.Reader
{
    /// <summary>
    /// Interface for readers that discover package-level information before document readers execute.
    /// </summary>
    internal interface IDiscoveryReader : IPluginReader
    {
        /// <summary>
        /// Gets or sets the reader options used for discovery validation.
        /// </summary>
        IOptions Options { get; set; }

        /// <summary>
        /// Initializes discovery with a caller-owned ZIP archive.
        /// </summary>
        void Init(ZipArchive archive, Workbook workbook, IOptions readerOptions);
    }
}
