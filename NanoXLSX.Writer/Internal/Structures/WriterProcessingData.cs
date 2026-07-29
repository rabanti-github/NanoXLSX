/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using NanoXLSX.Interfaces.Writer;
using NanoXLSX.Styles;

namespace NanoXLSX.Internal.Structures
{
    /// <summary>
    /// Class for data that is used in processors
    /// </summary>
    internal class WriterProcessingData : IWriterProcessingData
    {
        /// <summary>
        /// Style manager instance
        /// </summary>
        public StyleManager StyleManager { get; set; }

        /// <summary>
        /// Style repository instance
        /// </summary>
        public StyleRepository StyleRepository { get; set; }

        public WriterProcessingData(Workbook workbook, StyleRepository styleRepository)
        {
            this.StyleRepository = styleRepository;
            this.StyleManager = StyleManager.GetManagedStyles(workbook);
        }
    }
}
