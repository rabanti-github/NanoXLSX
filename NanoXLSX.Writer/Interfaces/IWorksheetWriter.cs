/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2025
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */


namespace NanoXLSX.Interfaces.Writer
{
    /// <summary>
    /// Interface, used by worksheet writers
    /// </summary>
    internal interface IWorksheetWriter : IPluginWriter
    {
        /// <summary>
        /// Gets the currently active worksheet
        /// </summary>
        Worksheet CurrentWorksheet { get; set; }
    }
}
