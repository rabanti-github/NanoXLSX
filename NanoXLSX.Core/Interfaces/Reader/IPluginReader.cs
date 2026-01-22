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
    /// Interface, used by XML reader classes 
    /// </summary>
    internal interface IPluginReader : IPlugin
    {
        /// <summary>
        /// Gets or replaces the workbook instance, defined by the constructor
        /// </summary>
        Workbook Workbook { get; set; }
    }
}
