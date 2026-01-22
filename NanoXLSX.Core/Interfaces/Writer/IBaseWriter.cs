/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2025
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */


using NanoXLSX.Styles;

namespace NanoXLSX.Interfaces.Writer
{
    /// <summary>
    /// Interface, used by writers
    /// </summary>
    internal interface IBaseWriter
    {
        /// <summary>
        /// Gets the workbook instance used by writer
        /// </summary>
        Workbook Workbook { get; }

        /// <summary>
        /// Gets the style manager instance used by writer
        /// </summary>
        StyleManager Styles { get; }

        /// <summary>
        /// Gets or set the writer to write shared strings
        /// </summary>
        ISharedStringWriter SharedStringWriter { get; set; }

    }
}
