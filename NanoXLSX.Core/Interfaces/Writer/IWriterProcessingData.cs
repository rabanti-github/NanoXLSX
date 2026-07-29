/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using NanoXLSX.Styles;

namespace NanoXLSX.Interfaces.Writer
{
    /// <summary>
    /// Interface, used by writing processors (preparation)
    /// </summary>
    internal interface IWriterProcessingData
    {
        /// <summary>
        /// Style manager instance, that can be accessed during the write preparation
        /// </summary>
        StyleManager StyleManager { get; set; }

        /// <summary>
        /// StyleRepository instance, that can be accessed during the write preparation
        /// </summary>
        StyleRepository StyleRepository { get; set; }

        // TODO add further relevant data for processing here - To be implemented in interface implementations
    }
}
